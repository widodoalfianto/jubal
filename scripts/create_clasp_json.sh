#!/usr/bin/env bash
# Create a clasp project settings file from .clasp.json.example
# Usage: SCRIPT_ID=yourScriptId ./scripts/create_clasp_json.sh
# Optional: CLASP_PROJECT_FILE=.clasp.prod.json
set -e

if [ -z "${SCRIPT_ID}" ]; then
  echo "Usage: SCRIPT_ID=yourScriptId $0"
  exit 2
fi

TARGET_FILE="${CLASP_PROJECT_FILE:-.clasp.json}"

if [ ! -f ".clasp.json.example" ]; then
  echo ".clasp.json.example not found in repo root."
  exit 3
fi

printf '{\n  "scriptId": "%s",\n  "rootDir": "./"\n}\n' "${SCRIPT_ID}" > "${TARGET_FILE}"
chmod 0644 "${TARGET_FILE}"

echo "Wrote ${TARGET_FILE} with scriptId=${SCRIPT_ID}. Review before committing."
exit 0
