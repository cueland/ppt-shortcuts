#!/usr/bin/env python3
"""
Generate docs/shortcuts.json from the key bank defined in docs/commands.js.

Why a bank: the add-in-only manifest fixes shortcuts at load time, and this PowerPoint
build lacks KeyboardShortcuts 1.1 (no runtime replaceShortcuts). So we register every
key we might ever want up front as a generic "slot", and a keymap stored in the add-in
decides which command each slot runs. Assigning a key never needs a redeploy.

The bank lives in commands.js (KEY_BANK_MODIFIER_SETS / KEY_BANK_KEYS) so the page is
self-contained; this script parses those two lines so the JSON can never drift from them.

Usage:
  scripts/build-shortcuts.py          # write docs/shortcuts.json
  scripts/build-shortcuts.py --check  # exit 1 if docs/shortcuts.json is out of date

After changing the bank: run this, bump ?v= on ExtendedOverrides in manifest.xml, push,
re-sideload.
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
COMMANDS_JS = os.path.join(DOCS, "commands.js")
SHORTCUTS_JSON = os.path.join(DOCS, "shortcuts.json")

# Action "name" must match the extended-manifest schema: ^[A-Za-z0-9-_+]+$ (no spaces).
NAME_PREFIX = "ChristiantialElements-"


def read_bank():
    src = open(COMMANDS_JS, encoding="utf-8").read()

    def const(name):
        m = re.search(rf"^const {name} = (\[.*?\]);", src, re.M)
        if not m:
            sys.exit(f"could not find `const {name} = [...]` on one line in commands.js")
        return json.loads(m.group(1))

    mods = const("KEY_BANK_MODIFIER_SETS")
    keys = const("KEY_BANK_KEYS")
    return [f"{m}+{k}" for m in mods for k in keys]


def slot_id(combo):
    return "k_" + combo.replace("+", "_")


def build(bank):
    for c in bank:
        name = NAME_PREFIX + c
        if not re.fullmatch(r"[A-Za-z0-9\-_+]+", name):
            sys.exit(f"action name {name!r} violates the schema pattern")
    return {
        "actions": [{"id": slot_id(c), "type": "ExecuteFunction", "name": NAME_PREFIX + c} for c in bank],
        "shortcuts": [{"action": slot_id(c), "key": {"default": c, "mac": c}} for c in bank],
    }


def main():
    bank = read_bank()
    data = build(bank)
    text = json.dumps(data, indent=2) + "\n"
    if "--check" in sys.argv:
        current = open(SHORTCUTS_JSON, encoding="utf-8").read() if os.path.exists(SHORTCUTS_JSON) else ""
        if current != text:
            sys.exit("docs/shortcuts.json is out of date — run scripts/build-shortcuts.py")
        print(f"shortcuts.json in sync ({len(bank)} slots)")
        return
    with open(SHORTCUTS_JSON, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"wrote {len(bank)} slots to docs/shortcuts.json")


if __name__ == "__main__":
    main()
