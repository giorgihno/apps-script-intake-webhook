const CONFIG = {
  INTAKE_SHEET_NAME: 'Intake',
  SUMMARY_SHEET_NAME: 'Summary',
  TIMEZONE: Session.getScriptTimeZone(), // uses account/script TZ
  RAW_PAYLOAD_MAX_CHARS: 5000,
  REQUIRED_FIELDS: ['firstName', 'lastName', 'email'],
};

function doPost(e) {
  const startedAt = new Date();
  const requestId = generateRequestId_();

  try {
    const contentType = (e && e.postData && e.postData.type) ? String(e.postData.type) : 'unknown';
    const rawBody = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';

    const parsed = parsePayload_(e, rawBody, contentType);
    const normalized = normalizePayload_(parsed);

    const validation = validatePayload_(normalized);
    if (!validation.ok) {
      appendRow_({
        requestId,
        contentType,
        rawBody,
        payload: normalized,
        status: 'rejected',
        notes: validation.message,
        startedAt,
      });

      console.log(`[${requestId}] Rejected: ${validation.message}`);
      return jsonResponse_(400, {
        status: 'error',
        requestId,
        message: validation.message,
      });
    }

    appendRow_({
      requestId,
      contentType,
      rawBody,
      payload: normalized,
      status: 'accepted',
      notes: '',
      startedAt,
    });

    

    console.log(`[${requestId}] Accepted: ${normalized.email}`);
    return jsonResponse_(200, {
      status: 'ok',
      requestId,
      message: 'Submission stored.',
    });

  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);

    // Best-effort logging row (don’t fail silently)
    try {
      const contentType = (e && e.postData && e.postData.type) ? String(e.postData.type) : 'unknown';
      const rawBody = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
      appendRow_({
        requestId,
        contentType,
        rawBody,
        payload: {},
        status: 'rejected',
        notes: `Exception: ${msg}`,
        startedAt,
      });
    } catch (logErr) {
      // Avoid cascading failures
      console.log(`[${requestId}] Failed to write error row: ${logErr}`);
    }

    console.log(`[${requestId}] Exception: ${msg}`);
    return jsonResponse_(500, {
      status: 'error',
      requestId,
      message: 'Internal error. See logs for details.',
    });
  }
}

/**
 * Parse request payload.
 * - JSON preferred if content type indicates JSON or body looks like JSON.
 * - Fallback to form-encoded using e.parameter 
 */
function parsePayload_(e, rawBody, contentType) {
  const ct = (contentType || '').toLowerCase().trim();

  // JSON path
  const looksJson = rawBody && rawBody.trim().startsWith('{');
  if (ct.includes('application/json') || looksJson) {
    if (!rawBody) return {};
    try {
      return JSON.parse(rawBody);
    } catch (err) {
      // Explicit malformed JSON error
      throw new Error('Malformed JSON payload. Ensure valid JSON body and correct Content-Type.');
    }
  }

  // Form-encoded fallback
  // e.parameter gives single values; e.parameters gives arrays
  const obj = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach((k) => {
      obj[k] = e.parameter[k];
    });
  }
  return obj;
}

// Basic required-field validation

function validatePayload_(p) {
  if (!p || Object.keys(p).length === 0) {
    return { ok: false, message: 'Empty payload.' };
  }

  const missing = [];
  CONFIG.REQUIRED_FIELDS.forEach((f) => {
    if (!p[f] || String(p[f]).trim() === '') missing.push(f);
  });

  if (missing.length) {
    return { ok: false, message: `Missing required field(s): ${missing.join(', ')}` };
  }

  // Minimal email sanity check (avoid over-rejecting)
  const email = String(p.email).trim();
  if (!email.includes('@') || email.startsWith('@') || email.endsWith('@')) {
    return { ok: false, message: 'Invalid email format.' };
  }

  return { ok: true, message: '' };
}

// Normalize incoming fields to a consistent schema
function normalizePayload_(p) {
  const out = {};

  // Accept common variants defensively
  out.firstName = pick_(p, ['firstName', 'first_name', 'firstname', 'FirstName', 'First Name']);
  out.lastName  = pick_(p, ['lastName', 'last_name', 'lastname', 'LastName', 'Last Name']);
  out.email     = pick_(p, ['email', 'Email', 'e-mail', 'mail']);
  out.phone     = pick_(p, ['phone', 'Phone', 'phoneNumber', 'phone_number']);

  out.source    = pick_(p, ['source', 'Source']) || 'unknown';
  out.formName  = pick_(p, ['formName', 'form_name', 'form', 'FormName']) || 'unknown';

  // Trim and standardize
  out.firstName = toTitleCase_(safeTrim_(out.firstName));
  out.lastName  = toTitleCase_(safeTrim_(out.lastName));
  out.email     = safeTrim_(out.email).toLowerCase();
  out.phone     = normalizePhone_(safeTrim_(out.phone));

  // Preserve meta if present (stringify later if needed)
  out.meta = p && p.meta ? p.meta : null;

  return out;
}

// Persist request to Intake sheet for audit/debugging

function appendRow_({ requestId, contentType, rawBody, payload, status, notes, startedAt }) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.INTAKE_SHEET_NAME);
  if (!sh) throw new Error(`Missing sheet tab: ${CONFIG.INTAKE_SHEET_NAME}`);

  const now = new Date();
  const utc = Utilities.formatDate(now, 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const local = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  const rawTrunc = truncate_(rawBody || JSON.stringify(payload || {}), CONFIG.RAW_PAYLOAD_MAX_CHARS);

  // IP and User-Agent are not reliably available in GAS Web Apps
  const ip = 'N/A';
  const userAgent = 'N/A';

  const row = [
    utc,
    local,
    requestId,
    contentType || 'unknown',
    ip,
    payload.source || 'unknown',
    payload.formName || 'unknown',
    payload.firstName || '',
    payload.lastName || '',
    payload.email || '',
    payload.phone || '',
    rawTrunc,
    userAgent,
    status,
    notes || '',
  ];

  sh.appendRow(row);

  const elapsedMs = new Date().getTime() - startedAt.getTime();
  Logger.log(`[${requestId}] Row appended. status=${status} elapsedMs=${elapsedMs}`);
}

// Build JSON response body
function jsonResponse_(statusCode, obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // Apps Script doesn't let you set HTTP status directly in all contexts.
  // Still include status in body and rely on consistent JSON response.
  // Some environments do respect this header via setHeader on HtmlService, but keep it simple.
  return output;
}

// Helpers

function pick_(obj, keys) {
  if (!obj) return '';
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') {
      return obj[k];
    }
  }
  return '';
}

function safeTrim_(v) {
  return (v === undefined || v === null) ? '' : String(v).trim();
}

function truncate_(s, maxChars) {
  const str = String(s || '');
  if (str.length <= maxChars) return str;
  return str.substring(0, maxChars) + '...<truncated>';
}

function toTitleCase_(s) {
  const str = String(s || '').trim();
  if (!str) return '';
  return str
    .toLowerCase()
    .split(' ')
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(' ');
}

// Normalize phone number for storage
function normalizePhone_(phone) {
  const p = String(phone || '').trim();
  if (!p) return '';
  const hasPlus = p.startsWith('+');
  const digits = p.replace(/[^\d]/g, '');
  if (!digits) return '';
  return hasPlus ? `+${digits}` : digits;
}

function generateRequestId_() {
  // Simple unique id suitable for logging/tracing
  const ts = Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMddHHmmss");
  const rand = Math.random().toString(36).substring(2, 8);
  return `req_${ts}_${rand}`;
}
