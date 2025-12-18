# Google Apps Script Intake Webhook

Google Apps Script Web App that receives POST requests, validates/normalizes the payload,
and appends results to a Google Sheet (Intake tab). Both accepted and rejected requests are logged.

## Key behavior
- Entry point: doPost(e)
- Payload types: application/json and application/x-www-form-urlencoded
- Required fields: firstName, lastName, email
- Normalization: trims names, lowercases email, normalizes phone

## Verification
Send a POST request to the deployed Web App URL and confirm a new row is appended in the Intake sheet.
