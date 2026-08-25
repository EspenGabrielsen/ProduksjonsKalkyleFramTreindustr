#!/usr/bin/env sh
set -eu

PORT="${PORT:-8000}"
BASE_URL="${MARIMO_BASE_URL:-}"

if [ -n "$BASE_URL" ]; then
  exec marimo run src/varekost_app.py --host 0.0.0.0 --port "$PORT" --headless --base-url "$BASE_URL"
fi

exec marimo run src/varekost_app.py --host 0.0.0.0 --port "$PORT" --headless