#!/usr/bin/env bash
# Deploys to Cloudflare Pages (https://grading-console.pages.dev).
# The whole site is inlined into one module worker, uploaded as Pages' _worker.js
# inside the deployment request, so no separate asset upload is needed.
# Requires CLOUDFLARE_ACCOUNT_ID and CLOUDFLARE_API_TOKEN (Pages: Edit).
set -euo pipefail
cd "$(dirname "$0")/.."
node deploy/build-worker.mjs
curl -fsS -X POST \
  -H "Authorization: Bearer ${CLOUDFLARE_API_TOKEN}" \
  "https://api.cloudflare.com/client/v4/accounts/${CLOUDFLARE_ACCOUNT_ID}/pages/projects/grading-console/deployments" \
  -F 'manifest={}' -F branch=main \
  -F '_worker.js=@deploy/out/worker.js;filename=_worker.js;type=application/javascript+module'
echo
