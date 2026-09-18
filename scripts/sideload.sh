#!/usr/bin/env bash
# Copy manifest.xml into PowerPoint's sideload directory (build brief §8), then restart PowerPoint.
# Note the container name: com.microsoft.Powerpoint — lowercase "p" in "point".
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

mkdir -p "$WEF"
# The filename is part of PowerPoint's registration of a sideloaded add-in: documents that
# used the add-in remember it, and if the file is renamed they show "This add-in is no
# longer available". So the file keeps its ORIGINAL name forever. One manifest only.
rm -f "$WEF/christiantial-elements.manifest.xml"
cp "$ROOT/manifest.xml" "$WEF/ppt-shortcuts.manifest.xml"
echo "Copied manifest to $WEF/ppt-shortcuts.manifest.xml"

if pgrep -xq "Microsoft PowerPoint"; then
  echo "Quitting PowerPoint…"
  osascript -e 'tell application "Microsoft PowerPoint" to quit'
  # wait for the process to exit
  for _ in $(seq 1 30); do pgrep -xq "Microsoft PowerPoint" || break; sleep 0.5; done
fi

echo "Starting PowerPoint…"
open -a "Microsoft PowerPoint"
echo "Done. Open a presentation; the add-in appears under Home › ChristiantialElements."
