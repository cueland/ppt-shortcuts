#!/usr/bin/env python3
"""
Generate manifest.xml: shared runtime, Home-tab "Open" button, and the ChristiantialElements
ribbon tab whose buttons run commands via ExecuteFunction → Office.actions.associate("ribbon_<id>").

Edit RIBBON below to change the tab, then:
  python3 scripts/build-manifest.py && scripts/sideload.sh
The ribbon is fixed in the manifest (Office has no runtime ribbon API for add-in-only manifests),
so every change here means a re-sideload. Icons are docs/assets/icons/<icon>-{16,32,80}.png,
rendered from the pane's SVG set (see README).
"""
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = "https://cueland.github.io/ppt-shortcuts"
ADDIN_ID = "4e27ba64-4cfa-4081-92ec-9daba7554361"
VERSION = "0.3.4.0"
SHORTCUTS_V = "7"   # bump when shortcuts.json changes

B = lambda cid, label, icon, desc: {"id": cid, "label": label, "icon": icon, "desc": desc}
M = lambda mid, label, icon, items, desc="", nolabel=False: {"menu": mid, "label": label, "icon": icon, "items": items, "desc": desc or label, "nolabel": nolabel}

RIBBON = [
    ("Position", "alignLeft", [
        B("alignLeft", "Left", "alignLeft", "Left edges to the reference; single shape → slide"),
        B("alignCenter", "Centre", "alignCenter", "Horizontal centres to the reference"),
        B("alignRight", "Right", "alignRight", "Right edges to the reference"),
        B("alignTop", "Top", "alignTop", "Top edges to the reference"),
        B("alignMiddle", "Middle", "alignMiddle", "Vertical centres to the reference"),
        B("alignBottom", "Bottom", "alignBottom", "Bottom edges to the reference"),
        B("dockLeft", "Dock L", "dockLeft", "Move left until touching the reference"),
        B("distributeH", "Distr. H", "distributeH", "Distribute horizontally: outer two stay, gaps evened"),
        B("dockRight", "Dock R", "dockRight", "Move right until touching the reference"),
        B("dockUp", "Dock U", "dockUp", "Move up until touching the reference"),
        B("distributeV", "Distr. V", "distributeV", "Distribute vertically: outer two stay, gaps evened"),
        B("dockDown", "Dock D", "dockDown", "Move down until touching the reference"),
        B("stackH", "Stack H", "stackH", "Butt together left to right, in selection order"),
        B("stackV", "Stack V", "stackV", "Butt together top to bottom, in selection order"),
        B("swap", "Swap", "swap", "Exchange two shapes' positions and layer order"),
        B("alignInTable", "In table", "table", "Snap loose shapes into the table cells they overlap"),
        B("matrix", "Align in matrix", "matrix", "Rows x cols from the pane settings"),
        B("goldenCanon", "Golden canon", "golden", "Inside the reference, bottom margin = 2x top"),
        M("nudge", "Nudge", "nudgeRight", [
            B("nudgeLeft", "Left", "nudgeLeft", "Move by the nudge amount"),
            B("nudgeRight", "Right", "nudgeRight", "Move by the nudge amount"),
            B("nudgeUp", "Up", "nudgeUp", "Move by the nudge amount"),
            B("nudgeDown", "Down", "nudgeDown", "Move by the nudge amount"),
        ], "Nudge by the amount set in the pane", nolabel=True),
    ]),
    ("Size", "matchWidth", [
        B("stretchLeft", "Stretch L", "stretchLeft", "Extend to the reference's far-left edge"),
        B("matchWidth", "Width", "matchWidth", "Target width := reference width"),
        B("stretchRight", "Stretch R", "stretchRight", "Extend to the reference's far-right edge"),
        B("stretchUp", "Stretch U", "stretchUp", "Extend to the reference's top edge"),
        B("matchHeight", "Height", "matchHeight", "Target height := reference height"),
        B("stretchDown", "Stretch D", "stretchDown", "Extend to the reference's bottom edge"),
        B("fillLeft", "Gap L", "fillLeft", "Grow left to touch the reference"),
        B("matchBoth", "Both", "matchBoth", "Both dimensions, non-proportional"),
        B("fillRight", "Gap R", "fillRight", "Grow right to touch the reference"),
        B("fillUp", "Gap U", "fillUp", "Grow up to touch the reference"),
        B("fitInside", "Fit inside", "fitInside", "Scale proportionally to fit within the reference"),
        B("fillDown", "Gap D", "fillDown", "Grow down to touch the reference"),
        B("fillOutside", "Fill ref", "fillOutside", "Scale proportionally to cover the reference"),
        B("resizeUp", "Bigger", "resizeUp", "Magic Resizer: scale by the factor in the pane"),
        B("resizeDown", "Smaller", "resizeDown", "Magic Resizer: scale by 1 / factor"),
        B("slice", "Slice / multiply", "slice", "Split one shape into rows x cols"),
    ]),
    ("Colour", "fill", [
        M("fillc", "Fill", "fill", [B(f"fill{i}", f"Slot {i}", "fill", f"Fill colour = palette slot {i}") for i in range(1, 11)], "Fill colour from the palette"),
        M("linec", "Line", "line", [B(f"line{i}", f"Slot {i}", "line", f"Line colour = palette slot {i}") for i in range(1, 11)], "Line colour from the palette"),
        M("fontc", "Font", "font", [B(f"font{i}", f"Slot {i}", "font", f"Font colour = palette slot {i}") for i in range(1, 11)], "Font colour from the palette"),
    ]),
    ("Text", "margins", [
        B("setMargins", "Margins", "margins", "Apply the margins from the pane"),
        B("marginsZero", "Zero margins", "marginsZero", "All four text margins to 0"),
        B("fitFormToText", "Fit to text", "fitText", "Resize the shape to its text"),
        B("wrapToggle", "Wrap", "wrap", "Toggle word wrap"),
        B("splitTextBox", "Split", "split", "Two boxes from one, at the cursor"),
        B("mergeTextBoxes", "Merge", "merge", "Combine text boxes in selection order"),
        B("bulletsToggle", "Bullets", "bullets", "Toggle bullets"),
        B("setFontSize", "Font size", "fontSize", "Apply the font size from the pane"),
    ]),
    ("Format", "pickup", [
        B("formatPickup", "Pick up", "pickup", "Pick up the last-selected shape's format; applies to the others if several are selected"),
        B("formatApply", "Apply", "apply", "Apply the picked-up format to the selection"),
        B("formatPainter", "Painter", "painter", "Apply the picked-up format to every new selection until toggled off"),
        B("saveMyFormat", "Save", "star", "Store the picked-up format as a named preset"),
        B("myFormat1", "My Format 1", "preset", "Apply saved format #1"),
        M("formats", "My Formats", "preset", [B(f"myFormat{i}", f"My Format {i}", "preset", f"Apply saved format #{i}") for i in range(2, 6)], "My Formats 2-5"),
    ]),
    ("Tools", "similar", [
        M("tools", "Tools", "similar", [
            B("selectSimilar", "Select similar", "similar", "Select shapes with the same type and fill"),
            B("decomposeTable", "Decompose table", "decompose", "Table → one text box per cell"),
            B("slidesAsPictures", "Slides as pictures", "pictures", "Selected slides tiled as images on a new slide"),
            B("exportImage", "Export as image", "export", "Render the selected shape to PNG in the pane"),
            B("hide", "Hide selected", "hide", "Hide (keeps position and layer)"),
            B("unhide", "Unhide all", "unhide", "Show every hidden shape on the slide"),
            B("goToSlide", "Go to slide", "goto", "Jump to the slide number in the pane"),
            B("agendaWizard", "Agenda", "agenda", "Agenda + divider slides from the pane's list"),
            B("masterLabelAdd", "Master label +", "master", "Add the label to every slide layout"),
            B("masterLabelRemove", "Master label -", "masterOff", "Remove it from every layout"),
        ]),
        B("addSticky", "Sticky", "sticky", "Reviewer note with initials + timestamp, in the default colour"),
        M("stickyc", "Colour", "sticky", [B(f"sticky_{n.lower()}", n, "sticky", f"Sticky in {n}") for n in ["Yellow", "Green", "Pink", "Orange", "Blue", "Purple"]], "Sticky in a specific colour"),
        B("togglePane", "Pane", "pane", "Show / hide the ChristiantialElements pane"),
    ]),
]

# ---------------------------------------------------------------------------
images, short, long_ = {}, {}, {}
def img(icon):
    for s in (16, 32, 80):
        images[f"Icon.{icon}.{s}"] = f"{BASE}/assets/icons/{icon}-{s}.png"
    return f'<Icon><bt:Image size="16" resid="Icon.{icon}.16"/><bt:Image size="32" resid="Icon.{icon}.32"/><bt:Image size="80" resid="Icon.{icon}.80"/></Icon>'
def sid(key, text):
    rid = "S." + re.sub(r"[^A-Za-z0-9]", "_", key)
    short[rid] = text[:125]
    return rid
def lid(key, text):
    rid = "L." + re.sub(r"[^A-Za-z0-9]", "_", key)
    long_[rid] = text[:250]
    return rid

def button_xml(b, as_item=False):
    tag = "Item" if as_item else 'Control xsi:type="Button"'
    close = "Item" if as_item else "Control"
    # Top-level buttons show only their icon (label = non-breaking space; the schema requires
    # a Label). Menu items keep their text. The tooltip still carries the real name.
    label_resid = sid('lbl.' + b['id'], b['label']) if as_item else "S.blank"
    return f'''<{tag} id="CE.{b["id"]}">
  <Label resid="{label_resid}"/>
  <Supertip><Title resid="{sid('lbl.' + b['id'], b['label'])}"/><Description resid="{lid('desc.' + b['id'], b['desc'])}"/></Supertip>
  {img(b["icon"])}
  <Action xsi:type="ExecuteFunction"><FunctionName>ribbon_{b["id"]}</FunctionName></Action>
</{close}>'''

def menu_xml(m):
    items = "\n".join(button_xml(b, as_item=True) for b in m["items"])
    label_resid = "S.blank" if m.get("nolabel") else sid('menu.' + m['menu'], m['label'])
    return f'''<Control xsi:type="Menu" id="CE.menu.{m["menu"]}">
  <Label resid="{label_resid}"/>
  <Supertip><Title resid="{sid('menu.' + m['menu'], m['label'])}"/><Description resid="{lid('menu.desc.' + m['menu'], m['desc'])}"/></Supertip>
  {img(m["icon"])}
  <Items>
{items}
  </Items>
</Control>'''

groups_xml = []
for gname, gicon, controls in RIBBON:
    cx = "\n".join(menu_xml(c) if "menu" in c else button_xml(c) for c in controls)
    groups_xml.append(f'''<Group id="CE.group.{gname}">
  <Label resid="{sid('group.' + gname, gname)}"/>
  {img(gicon)}
{cx}
</Group>''')

tab_xml = "\n".join(groups_xml)
open_icon = img("group")  # Home-tab button

def strings(d, tag):
    return "\n".join(f'        <bt:String id="{k}" DefaultValue="{v.replace("&", "&amp;").replace(chr(34), "&quot;").replace("<", "&lt;")}"/>' for k, v in d.items())
def image_res():
    return "\n".join(f'        <bt:Image id="{k}" DefaultValue="{v}"/>' for k, v in images.items())

short["S.blank"] = "\u00a0"
sid("tab", "ChristiantialElements")
sid("home.group", "ChristiantialElements"); sid("home.open", "Open"); lid("home.open.desc", "Open the ChristiantialElements pane: icons, key assignment, settings, log.")
sid("gs.title", "ChristiantialElements loaded"); lid("gs.desc", "Shortcuts are registered. Press Ctrl+Shift+Option+K or use the ChristiantialElements tab.")

manifest = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!-- GENERATED by scripts/build-manifest.py — edit that, not this. Sideloaded locally, NOT served. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>{ADDIN_ID}</Id>
  <Version>{VERSION}</Version>
  <ProviderName>Christian Ueland</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="ChristiantialElements"/>
  <Description DefaultValue="ChristiantialElements: keyboard-driven shape tools for PowerPoint on macOS."/>
  <IconUrl DefaultValue="{BASE}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{BASE}/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://github.com/cueland/ppt-shortcuts"/>
  <AppDomains>
    <AppDomain>https://cueland.github.io</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{BASE}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long"/>
        </Runtimes>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="S_gs_title"/>
            <Description resid="L_gs_desc"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CE.home">
                <Label resid="S_home_group"/>
                {open_icon}
                <Control xsi:type="Button" id="CE.home.open">
                  <Label resid="S_home_open"/>
                  <Supertip><Title resid="S_home_open"/><Description resid="L_home_open_desc"/></Supertip>
                  {open_icon}
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>PptShortcuts.Taskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
            <CustomTab id="CE.tab">
{tab_xml}
              <Label resid="S_tab"/>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{BASE}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{BASE}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{BASE}/assets/icon-80.png"/>
{image_res()}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/cueland/ppt-shortcuts"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{BASE}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
{strings(short, "short")}
      </bt:ShortStrings>
      <bt:LongStrings>
{strings(long_, "long")}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>

  <ExtendedOverrides Url="{BASE}/shortcuts.json?v={SHORTCUTS_V}"></ExtendedOverrides>
</OfficeApp>
'''
# resid values must be valid ids: replace the dots the sid/lid helpers produced
manifest = manifest.replace('resid="S.', 'resid="S_').replace('resid="L.', 'resid="L_')
manifest = re.sub(r'<bt:String id="S\.', '<bt:String id="S_', manifest)
manifest = re.sub(r'<bt:String id="L\.', '<bt:String id="L_', manifest)

out = os.path.join(ROOT, "manifest.xml")
with open(out, "w", encoding="utf-8") as f:
    f.write(manifest)
missing = [k for k, v in images.items() if not os.path.exists(os.path.join(ROOT, "docs", "assets", "icons", os.path.basename(v)))]
print(f"wrote manifest.xml: {sum(len(c) for _, _, c in RIBBON)} ribbon controls, {len(images)} image refs, {len(short)} short + {len(long_)} long strings")
if missing:
    print("MISSING ICONS:", missing)
