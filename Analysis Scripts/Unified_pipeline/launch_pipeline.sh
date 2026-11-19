#!/bin/bash
# Start the EMT Pipeline (Master Analysis Launcher) in the project's uv env.
#
#   ./launch_pipeline.sh
#
# "EMT Pipeline.app" calls this script too (build it with assets/make_app.sh).
# First-time setup: ./setup.sh

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Find uv. Apps started from Finder/Dock don't get your shell's PATH, so also
# check the usual install locations.
UV="$(command -v uv || true)"
if [ -z "$UV" ]; then
  for c in "$HOME/.local/bin/uv" "$HOME/.cargo/bin/uv" "/opt/homebrew/bin/uv" "/usr/local/bin/uv"; do
    if [ -x "$c" ]; then UV="$c"; break; fi
  done
fi
if [ -z "$UV" ]; then
  echo "Could not find uv. Install it (see README.md), then run ./setup.sh." >&2
  exit 1
fi

cd "$HERE"
# uv run creates/updates .venv from uv.lock if needed, then starts the launcher
exec "$UV" run --project "$HERE" python "$HERE/Python_analyze/Master_Analysis_Launcher.py" "$@"
