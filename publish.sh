#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

echo "Pushing to GitHub..."
git push -u origin main

echo "Enabling GitHub Pages (root of main)..."
if command -v gh >/dev/null 2>&1; then
  gh api -X POST "repos/sharmavaishaliweb-blip/information-Search-Survey/pages" \
    -f build_type=legacy \
    -f source='{"branch":"main","path":"/"}' \
    || gh api -X PUT "repos/sharmavaishaliweb-blip/information-Search-Survey/pages" \
      -f build_type=legacy \
      -f source='{"branch":"main","path":"/"}' \
    || true
  echo "Done. Live URL (may take 1–2 min):"
  echo "https://sharmavaishaliweb-blip.github.io/information-Search-Survey/"
else
  echo "gh CLI not found. After push, enable Pages manually:"
  echo "Repo → Settings → Pages → Source: Deploy from branch → main → / (root) → Save"
  echo "URL: https://sharmavaishaliweb-blip.github.io/information-Search-Survey/"
fi
