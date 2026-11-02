#!/usr/bin/env python3
"""Draws docs/images/how-it-works.svg, the "How it works" infographic for HVMTools.

    python3 tools/draw-how-it-works.py

Edit the text below and run it again; the layout (wrapping, card heights) is computed.
Adapted from tools/draw-how-it-works.py in com-morpheus-plugin (same author, MIT).
"""
from html import escape
from pathlib import Path

W = 1400
FONT = "-apple-system, 'Segoe UI', Helvetica, Arial, sans-serif"
INK, MUTED, LINE = "#1b1f24", "#57606a", "#d0d7de"
BLUE = ("#eaf2fc", "#2a6bb8")     # browser
TEAL = ("#ecf6f2", "#1c7a5c")     # Morpheus appliance / plugin
GREEN = ("#eef7e6", "#3f7a1d")    # HVM VMs
AMBER = ("#fff5e1", "#a35f00")    # guest agent check / gates
GRAY = ("#f6f8fa", "#57606a")

out: list[str] = []
MARKERS: set[str] = set()


def t(x, y, s, size=14, weight="normal", fill=INK, anchor="start", italic=False):
    st = ' font-style="italic"' if italic else ""
    out.append(f'<text x="{x}" y="{y}" font-size="{size}" font-weight="{weight}" fill="{fill}" '
               f'text-anchor="{anchor}"{st}>{escape(s)}</text>')


def rect(x, y, w, h, fill, stroke, rx=10, sw=1.5, dash=None):
    d = f' stroke-dasharray="{dash}"' if dash else ""
    out.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{d}/>')


def wrap(s, width, size):
    per = max(8, int(width / (size * 0.53)))
    words, lines, cur = s.split(), [], ""
    for w in words:
        if cur and len(cur) + 1 + len(w) > per:
            lines.append(cur)
            cur = w
        else:
            cur = f"{cur} {w}".strip()
    if cur:
        lines.append(cur)
    return lines


def para(x, y, w, s, size=13, fill=INK, gap=4, bullet=None, weight="normal"):
    """Writes wrapped text; returns the y below it."""
    indent = 14 if bullet else 0
    lines = wrap(s, w - indent, size)
    for i, ln in enumerate(lines):
        if i == 0 and bullet:
            t(x, y + size, bullet, size, "bold", fill)
        t(x + indent, y + size, ln, size, weight, fill)
        y += size + gap
    return y


def card(x, y, w, title, items, colors, size=13, title_size=15, bullet="•", draw=True):
    """A rounded card with a title and wrapped bullet items; returns its height."""
    global out
    saved = out
    out = []
    cy = y + 14
    t(x + 14, cy + title_size, title, title_size, "bold", colors[1])
    cy += title_size + 10
    for it in items:
        if isinstance(it, tuple):           # (text, style)
            text, style = it
            if style == "note":
                cy = para(x + 14, cy, w - 28, text, size - 1, MUTED) + 4
            elif style == "plain":
                cy = para(x + 14, cy, w - 28, text, size) + 4
            elif style == "code":
                cy = para(x + 14, cy, w - 28, text, size - 1, colors[1], weight="bold") + 4
            else:                           # custom bullet
                cy = para(x + 14, cy, w - 28, text, size, bullet=style) + 4
        else:
            cy = para(x + 14, cy, w - 28, it, size, bullet=bullet) + 4
    h = cy - y + 8
    body = out
    out = saved
    if draw:
        rect(x, y, w, h, "#ffffff", colors[1], rx=8, sw=1.2)
        out.extend(body)
    return h


def arrow(x1, y1, x2, y2, color=INK, label=None, lx=None, ly=None, both=False):
    mid = "a" + color.lstrip("#")
    MARKERS.add(color)
    m = f' marker-start="url(#{mid})"' if both else ""
    out.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="2" marker-end="url(#{mid})"{m}/>')
    if label:
        for i, ln in enumerate(label.split("\n")):
            t(lx if lx is not None else (x1 + x2) / 2, (ly if ly is not None else (y1 + y2) / 2 - 8) + i * 15,
              ln, 12, "bold", color, "middle")


def zone(x, y, w, h, title, sub, colors):
    rect(x, y, w, h, colors[0], colors[1], rx=14, sw=2)
    t(x + 18, y + 30, title, 19, "bold", colors[1])
    t(x + 18, y + 50, sub, 13, "normal", MUTED)




def stack(x, y, w, cards, gap=12):
    for title, items, colors in cards:
        y += card(x, y, w, title, items, colors) + gap
    return y


# ---------------------------------------------------------------- architecture
TOP = 108
COL1_X, COL1_W = 24, 306
COL2_X, COL2_W = 400, 600
COL3_X, COL3_W = 1070, 306

col1 = [
    ("Run one command", [
        ("python hvmtools.py --url https://<manager> -u <user>", "code"),
        "Password prompt, or --token, or MORPHEUS_* environment variables",
        ("Python 3.8+ with requests and openpyxl. Any machine that can reach the manager.", "note"),
    ], BLUE),
    ("Options", [
        "--no-snapshots · --no-stats · --raw · --snapshot-age · --workers",
        "--dump-json saves what was collected; --from-json rebuilds without a manager",
    ], BLUE),
]
col2_left = [
    ("Sign-in", [
        ("POST /oauth/token (client morph-api)", "code"),
        "Or a token you already have. After sign-in, GET only",
    ], TEAL),
    ("Inventory", [
        ("/api/ping · whoami · zones · groups · clusters · servers · instances · networks", "code"),
        "Paged with max and offset until every record is read",
    ], TEAL),
]
col2_right = [
    ("Storage and details", [
        ("/api/clusters/{id}/datastores and /api/zones/{id}/data-stores", "code"),
        "Usage stats from /api/servers/{id} only where the list lacks them",
    ], TEAL),
    ("Snapshots", [
        ("/api/instances/{id}/snapshots", "code"),
        "One call per instance, 8 in parallel",
    ], TEAL),
]
col3 = [
    ("One .xlsx, RVTools layout", [
        "vInfo · vCPU · vMemory · vDisk · vNetwork · vSnapshot",
        "vHost · vCluster · vDatastore · vNetworks · vInstance",
        "vHealth · vSource (+ raw_* tabs with --raw)",
    ], GREEN),
    ("Ready to use", [
        "Frozen headers and filters on every tab",
        "Sizes in MiB, dates in UTC",
        "vHealth rows shaded by severity",
    ], GREEN),
]


def measure(cards, w):
    return sum(card(0, 0, w, a, b, c, draw=False) + 12 for a, b, c in cards)


inner = 2 * 14
half = (COL2_W - 3 * 14) // 2
h1 = measure(col1, COL1_W - inner)
h2 = max(measure(col2_left, half), measure(col2_right, half))
h3 = measure(col3, COL3_W - inner)
DATA_H = 96
ZONE_H = 66 + max(h1 + 54, h2 + DATA_H + 18, h3 + 54)

zone(COL1_X, TOP, COL1_W, ZONE_H, "① Your computer", "hvmtools.py: one Python file", BLUE)
zone(COL2_X, TOP, COL2_W, ZONE_H, "② HPE VM Essentials manager", "Morpheus REST API: read-only calls", TEAL)
zone(COL3_X, TOP, COL3_W, ZONE_H, "③ Excel workbook", "HVMTools_<manager>_<date>.xlsx", GREEN)

stack(COL1_X + 14, TOP + 66, COL1_W - inner, col1)
stack(COL2_X + 14, TOP + 66, half, col2_left)
stack(COL2_X + 28 + half, TOP + 66, half, col2_right)
stack(COL3_X + 14, TOP + 66, COL3_W - inner, col3)

dy = TOP + ZONE_H - DATA_H - 14
rect(COL2_X + 14, dy, COL2_W - 28, DATA_H, "#ffffff", GRAY[1], rx=8, sw=1.2, dash="5 4")
t(COL2_X + 28, dy + 24, "Errors don't stop the run", 15, "bold", GRAY[1])
para(COL2_X + 28, dy + 34, COL2_W - 56,
     "A 403 on one endpoint (a missing permission) prints a warning and leaves that data empty. "
     "Only a 401 (bad credentials or token) stops the run. Credential-like fields are blanked before anything is saved.", 13)

para(COL1_X + 14, TOP + ZONE_H - 46, COL1_W - 28,
     "A read-only Morpheus user is enough.", 12, MUTED)
para(COL3_X + 14, TOP + ZONE_H - 46, COL3_W - 28,
     "Hand it to a customer or colleague, or feed it to a sizing tool.", 12, MUTED)

ay = TOP + 140
arrow(COL1_X + COL1_W + 4, ay, COL2_X - 4, ay, BLUE[1], "HTTPS\nGET only", ly=ay - 26, both=True)
arrow(COL2_X + COL2_W + 4, ay, COL3_X - 4, ay, GREEN[1], "rows per\ntab", ly=ay - 26)

# ---------------------------------------------------------------- the stages
FY = TOP + ZONE_H + 44
t(W / 2, FY, "What happens, step by step", 22, "bold", INK, "middle")
FY += 22
GAP = 16
PW = (W - 2 * 24 - 3 * GAP) / 4
flows = [
    ("A.  Collect  (collect())", TEAL, [
        ("Sign in, or use the token.", "1"),
        ("Read every list endpoint, page by page, into one dictionary.", "2"),
        ("Fill in missing usage stats and snapshots with per-item calls.", "3"),
        ("Progress goes to stderr, one line per step.", "note"),
    ]),
    ("B.  Match  (build_sheets())", BLUE, [
        ("Hosts are hypervisor servers; every other server is a VM.", "1"),
        ("VM → host → cluster; VM → instance; disk → datastore; NIC → network.", "2"),
        ("One list of rows per tab.", "3"),
        ("Works only on the collected data, so --from-json gives the same workbook offline.", "note"),
    ]),
    ("C.  Check  (vHealth)", AMBER, [
        ("Hosts in a bad state or maintenance mode; host memory above 90%", "•"),
        ("Datastores under 10% free", "•"),
        ("VMs in a bad state or powered off; running VMs without the Morpheus agent", "•"),
        ("Snapshots older than --snapshot-age days", "•"),
    ]),
    ("D.  Write  (write_workbook())", GREEN, [
        ("Formats every tab: Arial 10, dark-blue headers, freeze B2, autofilter, auto widths.", "1"),
        ("Empty tabs say \"No data returned\".", "2"),
        ("Saves one .xlsx and prints the row count per tab.", "3"),
    ]),
]


def flow_items(items):
    return [(s, b if b in ("plain", "note") else (f"{b}." if b.isdigit() else b)) for s, b in items]


fh = max(card(0, 0, PW, title, flow_items(items), col, size=13, title_size=15, draw=False)
         for title, col, items in flows)
for i, (title, col, items) in enumerate(flows):
    x = 24 + i * (PW + GAP)
    rect(x, FY, PW, fh, col[0], col[1], rx=12, sw=2)
    saved = len(out)
    card(x, FY, PW, title, flow_items(items), col, size=13, title_size=15)
    del out[saved]              # drop the white rect, keep the text

FOOT = FY + fh + 30
t(24, FOOT, "Unofficial community tool, MIT license. Not an HPE product. HPE, Morpheus and VM Essentials are trademarks "
  "of Hewlett Packard Enterprise; RVTools is a trademark of its owner.", 12, "normal", MUTED)
H = FOOT + 22

markers = "\n".join(f'<marker id="a{c.lstrip("#")}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{c}"/></marker>' for c in sorted(MARKERS))
head = f'''<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" font-family="{FONT}">
<title>How HVMTools works</title>
<desc>HVMTools reads an HPE VM Essentials manager through the Morpheus REST API (GET only) and writes an RVTools-style Excel workbook.</desc>
<defs>
{markers}
</defs>
<rect width="{W}" height="{H}" fill="#ffffff"/>
<text x="24" y="48" font-size="30" font-weight="bold" fill="{INK}">How HVMTools works</text>
<text x="24" y="76" font-size="15" fill="{MUTED}">RVTools-style inventory for HPE VM Essentials: one command reads the manager's REST API and writes one Excel workbook. Read-only.</text>
'''
out_path = Path(__file__).resolve().parent.parent / "docs/images/how-it-works.svg"
out_path.write_text(head + "\n".join(out) + "\n</svg>\n")
print(out_path, W, H)
