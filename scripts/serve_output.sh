#!/usr/bin/env bash
set -euo pipefail

OUTPUT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../output" && pwd)"

if python - <<'PY'
import sys
try:
    import http.server  # noqa: F401
except Exception:
    sys.exit(1)
PY
then
  python -m http.server 8000 --directory "${OUTPUT_DIR}"
else
  python -m SimpleHTTPServer 8000
fi
