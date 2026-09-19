#!/usr/bin/env python3
"""
Repair a deck whose stored Expropriated Elements auto-start reference names the wrong version.

  python3 scripts/fix-deck-reference.py "deck.pptx"        # writes deck.fixed.pptx next to it

Rewrites <we:reference id="<our guid>" version="…"/> in ppt/webextensions/*.xml to the manifest's
version (0.1.0.0). Everything else in the package is copied byte-for-byte.
"""
import re, shutil, sys, zipfile

GUID = "4e27ba64-4cfa-4081-92ec-9daba7554361"
VERSION = "0.1.0.0"

src = sys.argv[1]
dst = re.sub(r"\.pptx$", "", src) + ".fixed.pptx"
zin = zipfile.ZipFile(src)
changed = 0
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("ppt/webextensions/") and item.filename.endswith(".xml"):
            text = data.decode("utf-8")
            new = re.sub(r'(<we:reference[^>]*id="' + GUID + r'"[^>]*version=")[^"]*(")', r"\g<1>" + VERSION + r"\2", text)
            if new != text:
                changed += 1
                data = new.encode("utf-8")
        zout.writestr(item, data)
print(f"wrote {dst}: {changed} reference(s) set to {VERSION}")
