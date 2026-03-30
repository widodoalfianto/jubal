CI/CD for Jubal (Apps Script)

This document explains how the GitHub Actions workflow deploys the same Apps Script codebase to multiple tenant spreadsheets using `clasp`.

Prerequisites
- A Google Cloud project with the Apps Script API enabled and Drive API enabled.
- A Google Service Account with a JSON key that has permission to access the Apps Script project (Editor or owner on the target script) and Drive access for moving files.
- GitHub repository admin rights to add repository secrets.
- `appsscript.json` manifest present in the repo and the project laid out for `clasp`.
- `.claspignore` present so helper files do not get deployed.
- `deploy/tenants.json` present and committed.

Required repository secret
- `GCP_CREDENTIALS`: the full JSON contents of the service account key file.
  - Add it under GitHub repo Settings → Secrets → Actions → New repository secret.
  - Or set via `gh` CLI: `gh secret set GCP_CREDENTIALS --body "$(cat path/to/key.json)"`

Script ID options
- Each tenant in `deploy/tenants.json` can define either:
  - `scriptId`: the literal Apps Script project ID
  - `scriptIdVar`: the name of a GitHub variable or secret containing the script ID
- Your current setup uses GitHub variables for script IDs.
- Add these repository variables or environment variables:
  - `SCRIPT_ID_DEV_FULL`
  - `SCRIPT_ID_DEV_MINIMAL`
  - `SCRIPT_ID_OC`
  - `SCRIPT_ID_MRV`

Service Account setup (high level)
1. Create or pick a Google Cloud project.
2. Enable "Apps Script API" and "Google Drive API".
3. Create a Service Account (IAM & Admin → Service Accounts).
4. Grant the service account a role that allows updating Apps Script projects (e.g., Project Editor). If you want more restrictive scope, ensure it can use Apps Script API and read/write Drive files created by the script.
5. Create and download a JSON key for the service account.
6. Store the file contents in the `GCP_CREDENTIALS` GitHub secret.

Notes about `clasp` authentication
- The workflow uses `clasp login --creds gcp-creds.json` to authenticate with the service account JSON.
- For first-time setup you can also test locally:

```bash
npm install -g @google/clasp
clasp login
```

Ensure your local `clasp` project has an `appsscript.json` manifest and (optionally) `.clasp.json` with `scriptId`.

Tenant registry
- `deploy/tenants.json` is the source of truth for the tenant fleet.
- The default file includes:
  - `dev-full`
  - `dev-minimal`
  - `my-church`
  - `friend-church`
- Each tenant includes:
  - `name`
  - `stage` (`dev` or `prod`)
  - `scriptId` or `scriptIdVar`
  - `runRemoteTests`

Example:

```json
{
  "name": "my-church",
  "stage": "prod",
  "scriptIdVar": "SCRIPT_ID_MY_CHURCH",
  "runRemoteTests": false
}
```

Workflow behavior
- On push to `main`:
  - validate the repo layout and tenant registry
  - build the `dev` tenant matrix from `deploy/tenants.json`
  - deploy to all `dev` tenants
  - run `runAllTests()` remotely against tenants with `runRemoteTests: true`
- On manual `workflow_dispatch` with `deploy_stage=dev`:
  - deploy to `dev` tenants and run remote tests
- On manual `workflow_dispatch` with `deploy_stage=prod`:
  - deploy to all `prod` tenants in the registry
- If `clasp deploy` is not configured or fails, the workflow still pushes files with `clasp push`.

Checklist before running the workflow
- Confirm the repo contains `appsscript.json` (Apps Script manifest).
- Confirm `.claspignore` exists and only whitelists Apps Script files.
- Confirm `deploy/tenants.json` contains the tenants you want to manage.
- Add `GCP_CREDENTIALS` to GitHub secrets.
- Add the `SCRIPT_ID_*` GitHub variables or secrets referenced by each tenant.
- Share each target Apps Script project and spreadsheet with the deployment identity.
- For safety, configure GitHub environments:
  - `dev`
  - `prod`
  - Use required reviewers on `prod` before allowing production deploys.
- Add `GCP_CREDENTIALS` to GitHub secrets.

Troubleshooting
- If `clasp login --creds` fails, validate the JSON key and that the service account has necessary permissions.
- If a tenant deploy fails before `clasp push`, verify the tenant has either a valid `scriptId` or a matching GitHub variable/secret named by `scriptIdVar`.
- If `clasp deploy` fails, inspect the `appsscript.json` manifest and ensure deployment config exists; you may need to create a deployment manually one time via `clasp deploy` locally.
- If remote tests fail, verify:
  - the service account can execute the target script
  - the target spreadsheet is a non-production dev tenant
  - `Testing.js` is included in `.claspignore`

Security notes
- Grant the service account the least privilege required.
- Rotate keys periodically and remove unused service account keys.

Advanced
- You can extend the workflow to run tests (e.g., run unit test script that writes results to `Test Results` sheet) by adding a step to call a test runner using `clasp` remote execution or by running a local Node-based test harness if you add one.

Contact
- If you want, I can add a helper script or `Makefile` to locally prepare and validate the `clasp` setup. Ask me to add it and I will create `scripts/` helpers.

Local test & run commands
-------------------------

Quick local commands to validate and run the test helper added to this repo:

1) Validate environment and clasp setup:

```bash
bash scripts/validate_clasp.sh
```

2) Validate the tenant registry directly:

```bash
npm run validate:tenants
```

3) Create `.clasp.json` from a tenant script ID:

```bash
SCRIPT_ID=YOUR_SCRIPT_ID_HERE bash scripts/create_clasp_json.sh
```

You can also write a named target file:

```bash
SCRIPT_ID=YOUR_SCRIPT_ID_HERE CLASP_PROJECT_FILE=.clasp.dev-full.json bash scripts/create_clasp_json.sh
```

4) Run the Apps Script `runAllTests` remotely using the service account JSON:

```bash
export GCP_CREDENTIALS="$(cat path/to/key.json)"
export SCRIPT_ID=YOUR_SCRIPT_ID_HERE
npm install
node scripts/run_tests.js
```

Or using a credentials file path:

```bash
GCP_CREDENTIALS_FILE=path/to/key.json SCRIPT_ID=YOUR_SCRIPT_ID_HERE node scripts/run_tests.js
```

Notes:
- Ensure the service account email is granted Editor access to the Apps Script project and any spreadsheets the tests touch.
- The `run_tests` helper calls `runAllTests` in the Apps Script project (devMode=true) and prints the execution response.
- Only point remote tests at `dev` tenants or spreadsheet copies. `runAllTests()` modifies sheets and forms.
