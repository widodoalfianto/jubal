#!/usr/bin/env bash
# Validate local environment for clasp/Appsscript deploy
set -e

echo "Validating local environment for clasp..."

which node >/dev/null 2>&1 || { echo "node not found. Install Node.js (https://nodejs.org/)"; exit 2; }
which npm >/dev/null 2>&1 || { echo "npm not found. Install npm."; exit 2; }

if ! command -v clasp >/dev/null 2>&1; then
  echo "clasp not found. Installing @google/clasp globally..."
  npm install -g @google/clasp
fi

echo "clasp version: $(clasp --version 2>/dev/null || echo 'unknown')"

# Check for appsscript.json
if [ -f "appsscript.json" ]; then
  echo "Found appsscript.json"
else
  echo "Warning: appsscript.json not found in repo root. Create it or ensure your project layout is correct."
fi

# Check for .claspignore
if [ -f ".claspignore" ]; then
  echo "Found .claspignore"
else
  echo "Warning: .claspignore not found. Helper/test .js files may be pushed to Apps Script."
fi

# Check tenant registry
if [ -f "deploy/tenants.json" ]; then
  echo "Found deploy/tenants.json"
  if [ -f "scripts/tenant_matrix.js" ]; then
    node scripts/tenant_matrix.js --validate
  fi
else
  echo "Warning: deploy/tenants.json not found. Multi-tenant CI deploys will not work."
fi

# Check for .clasp.json or example
if [ -f ".clasp.json" ]; then
  echo "Found .clasp.json"
else
  if [ -f ".clasp.json.example" ]; then
    echo "No .clasp.json found, but .clasp.json.example exists. Run scripts/create_clasp_json.sh or copy it to .clasp.json and set scriptId."
  else
    echo "No .clasp.json or example found. Create .clasp.json with your scriptId."
  fi
fi

# If .clasp.json exists, show scriptId
if [ -f ".clasp.json" ]; then
  echo ".clasp.json contents:" && cat .clasp.json
  if grep -q "scriptId" .clasp.json; then
    echo "Attempting to show clasp status (does not require auth for reading local manifest)"
    clasp status || echo "clasp status failed (may require authentication)"
  fi
fi

echo "Validation complete. To authenticate clasp locally, run:"
echo "  clasp login"
echo "Or, if you have an OAuth desktop client JSON for clasp:"
echo "  clasp login --creds path/to/oauth-client.json"
echo "For CI or service-account based automation, prefer your workflow-specific auth path."

exit 0
