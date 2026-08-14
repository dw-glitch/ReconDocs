#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${project_root}"

echo "Running Next.js build (local/Vercel)..."
exec "${project_root}/node_modules/.bin/next" build
