#!/bin/bash
# One-time setup on macOS (re-run after a `git pull` that changes uv.lock or
# R_reports/renv.lock):
#   1. Python env + Quarto from uv.lock   (uv sync)
#   2. R packages from R_reports/renv.lock (renv::restore)
#   3. Optionally build "EMT Pipeline.app"

set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HERE"

UV="$(command -v uv || true)"
[ -z "$UV" ] && [ -x "$HOME/.local/bin/uv" ] && UV="$HOME/.local/bin/uv"
if [ -z "$UV" ]; then
  echo "uv is not installed. Install it with:" >&2
  echo "  curl -LsSf https://astral.sh/uv/install.sh | sh" >&2
  echo "then open a new terminal and run ./setup.sh again." >&2
  exit 1
fi

echo "== Python env and Quarto (uv sync) =="
"$UV" sync

RSCRIPT="$(command -v Rscript || true)"
for c in /usr/local/bin/Rscript /opt/homebrew/bin/Rscript /Library/Frameworks/R.framework/Resources/bin/Rscript; do
  [ -z "$RSCRIPT" ] && [ -x "$c" ] && RSCRIPT="$c"
done
if [ -z "$RSCRIPT" ]; then
  echo
  echo "R is not installed, so the reports can't run (the analyses still work)." >&2
  echo "Install R 4.5 from https://cran.r-project.org/bin/macosx/ and run ./setup.sh again." >&2
  exit 1
fi

echo
echo "== R packages (renv::restore, first time takes a few minutes) =="
cd "$HERE/R_reports"
"$RSCRIPT" -e 'renv::restore(prompt = FALSE)'
"$RSCRIPT" -e 'renv::status()'
cd "$HERE"

echo
if [ ! -d "$HERE/EMT Pipeline.app" ] && [ -t 0 ]; then
  read -r -p "Build the double-clickable 'EMT Pipeline.app'? [Y/n] " ans
  if [[ ! "$ans" =~ ^[Nn] ]]; then "$HERE/assets/make_app.sh"; fi
fi

echo
echo "Setup done. Start the pipeline with ./launch_pipeline.sh or 'EMT Pipeline.app'."
