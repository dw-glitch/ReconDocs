#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${project_root}"

if [[ "${VERCEL:-}" == "1" || -n "${VERCEL_ENV:-}" ]]; then
  echo "Running Vercel-compatible Next.js build..."
  exec "${project_root}/node_modules/.bin/next" build
fi

exec bash "${project_root}/scripts/build-verified.sh"
