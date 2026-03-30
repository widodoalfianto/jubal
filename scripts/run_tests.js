#!/usr/bin/env node
/*
  Run `runAllTests` in the Apps Script project remotely using a service account.

  Requirements:
  - The service account JSON must be provided via the env var `GCP_CREDENTIALS` (raw JSON string)
    or as a file path in `GCP_CREDENTIALS_FILE`.
  - The Apps Script project must be shared with the service account email (Editor access),
    or the script must be accessible to that service account.
  - Provide the `SCRIPT_ID` either via env var or `--scriptId` CLI arg.

  Example:
    export GCP_CREDENTIALS="$(cat path/to/key.json)"
    export SCRIPT_ID=1aBcDefGhI_jKl
    node scripts/run_tests.js

  Or:
    GCP_CREDENTIALS_FILE=path/to/key.json node scripts/run_tests.js --scriptId=1aBcDefGhI_jKl
*/

const { google } = require('googleapis');
const fs = require('fs');

function getArg(name) {
  const idx = process.argv.indexOf(name);
  if (idx >= 0 && idx + 1 < process.argv.length) return process.argv[idx + 1];
  return null;
}

function loadCredentials() {
  if (process.env.GCP_CREDENTIALS) {
    try {
      return JSON.parse(process.env.GCP_CREDENTIALS);
    } catch (e) {
      console.error('Failed to parse GCP_CREDENTIALS env var: ' + e.message);
      process.exit(2);
    }
  }
  if (process.env.GCP_CREDENTIALS_FILE) {
    try {
      const raw = fs.readFileSync(process.env.GCP_CREDENTIALS_FILE, 'utf8');
      return JSON.parse(raw);
    } catch (e) {
      console.error('Failed to read/parse GCP_CREDENTIALS_FILE: ' + e.message);
      process.exit(3);
    }
  }
  console.error('No GCP_CREDENTIALS or GCP_CREDENTIALS_FILE provided.');
  process.exit(4);
}

async function main() {
  const creds = loadCredentials();
  const scriptId = process.env.SCRIPT_ID || getArg('--scriptId') || getArg('-s');
  if (!scriptId) {
    console.error('Script ID not provided. Set SCRIPT_ID env var or pass --scriptId');
    process.exit(5);
  }

  // Scopes needed for running scripts and reading/writing Drive/Sheets if tests interact with them
  const scopes = [
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
  ];

  const jwtClient = new google.auth.JWT(
    creds.client_email,
    null,
    creds.private_key,
    scopes,
    null
  );

  try {
    await jwtClient.authorize();
  } catch (e) {
    console.error('Service account authorization failed: ' + e.message);
    console.error('Ensure the service account JSON is valid and the service account has access to the Apps Script project.');
    process.exit(6);
  }

  const service = google.script({ version: 'v1', auth: jwtClient });

  try {
    console.log('Calling Apps Script `runAllTests` on scriptId=' + scriptId + ' (devMode=true)');
    const res = await service.scripts.run({
      scriptId,
      requestBody: {
        function: 'runAllTests',
        parameters: [],
        devMode: true
      }
    });

    if (res.data && res.data.error) {
      console.error('Execution returned an error:');
      console.error(JSON.stringify(res.data.error, null, 2));
      process.exit(7);
    }

    console.log('Execution response:');
    console.log(JSON.stringify(res.data, null, 2));

    // If the Apps Script function returns a structured result, display it
    if (res.data && res.data.response && res.data.response.result) {
      console.log('Result:');
      console.log(JSON.stringify(res.data.response.result, null, 2));
    }

    process.exit(0);
  } catch (err) {
    console.error('API call failed: ' + (err && err.message ? err.message : String(err)));
    process.exit(8);
  }
}

main();
