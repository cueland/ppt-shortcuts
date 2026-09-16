#!/usr/bin/env bash
# Copy manifest.xml into PowerPoint's sideload directory (build brief §8), then restart PowerPoint.
# Note the container name: com.microsoft.Powerpoint — lowercase "p" in "point".
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

mkdir -p "$WEF"
rm -f "$WEF/ppt-shortcuts.manifest.xml"
cp "$ROOT/manifest.xml" "$WEF/christiantial-elements.manifest.xml"
echo "Copied manifest to $WEF/christiantial-elements.manifest.xml"

if pgrep -xq "Microsoft PowerPoint"; then
  echo "Quitting PowerPoint…"
  osascript -e 'tell application "Microsoft PowerPoint" to quit'
  # wait for the process to exit
  for _ in $(seq 1 30); do pgrep -xq "Microsoft PowerPoint" || break; sleep 0.5; done
fi

echo "Starting PowerPoint…"
open -a "Microsoft PowerPoint"
echo "Done. Open a presentation; the add-in appears under Home › ChristiantialElements."
