#!/usr/bin/env bash
# Clear PowerPoint's add-in web cache (build brief §9). Run when a pushed change doesn't show up.
# Quits PowerPoint first; the cache cannot be cleared reliably while it's running.
set -euo pipefail

CACHE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches"

if pgrep -xq "Microsoft PowerPoint"; then
  echo "Quitting PowerPoint…"
  osascript -e 'tell application "Microsoft PowerPoint" to quit'
  for _ in $(seq 1 30); do pgrep -xq "Microsoft PowerPoint" || break; sleep 0.5; done
fi

if [ -d "$CACHE" ]; then
  echo "Clearing $CACHE"
  rm -rf "${CACHE:?}"/*
else
  echo "No cache directory at $CACHE (nothing to clear)."
fi

# Office also keeps a shared WebKit cache for add-in web content on some builds.
WEF_CACHE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
if [ -d "$WEF_CACHE" ]; then
  echo "Clearing $WEF_CACHE"
  rm -rf "${WEF_CACHE:?}"/*
fi

echo "Done. Start PowerPoint again (or run scripts/sideload.sh)."
