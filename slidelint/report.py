"""SlideLint HTML report — score, per-check counts, slide-level findings."""

from __future__ import annotations

import html
from datetime import datetime
from typing import Dict, List

from .audit import summary_rows

CSS = """
:root{--bg:#0b0f18;--panel:#121a29;--line:#22304a;--fg:#e8eef8;--muted:#8b9ab1;
--accent:#38bdf8;--good:#34d399;--warn:#fbbf24;--bad:#f87171}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);
font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
.wrap{max-width:1060px;margin:0 auto;padding:34px 24px 60px}
h1{font-size:26px;margin:0 0 6px;letter-spacing:-.02em}
h2{font-size:14px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin:32px 0 10px}
.sub{color:var(--muted);font-size:13px}
.hero{display:flex;gap:20px;align-items:center;background:var(--panel);border:1px solid var(--line);
border-radius:14px;padding:20px 22px;margin-top:18px}
.hero .score{font-size:40px;font-weight:700;letter-spacing:-.03em;line-height:1}
.hero .vt{font-size:15.5px;font-weight:650}
.hero .vd{color:var(--muted);font-size:12.5px;margin-top:4px}
.meta{display:flex;flex-wrap:wrap;gap:8px 20px;font-size:12px;color:var(--muted);margin-top:12px}
.meta b{color:var(--fg)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-top:14px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:11px;padding:13px 15px}
.card .l{font-size:10.5px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted)}
.card .v{font-size:22px;font-weight:650;margin-top:5px}
.card .f{font-size:10.5px;color:var(--muted);margin-top:2px}
table{width:100%;border-collapse:collapse;margin-top:6px;font-size:12.5px}
th{text-align:left;color:var(--muted);font-weight:600;font-size:10.5px;text-transform:uppercase;
letter-spacing:.06em;border-bottom:1px solid var(--line);padding:7px 9px}
td{padding:7px 9px;border-bottom:1px solid #1a2436;vertical-align:top}
.badge{display:inline-block;padding:1px 8px;border-radius:999px;font-size:10.5px;
text-transform:uppercase;letter-spacing:.06em}
.badge.high{background:rgba(248,113,113,.14);color:#f87171}
.badge.medium{background:rgba(251,191,36,.14);color:#fbbf24}
.badge.low{background:rgba(139,154,177,.14);color:#8b9ab1}
.mono{font-family:ui-monospace,Menlo,monospace;font-size:11.5px;color:#cfe3f7}
.lock{background:linear-gradient(135deg,#131c2c,#0d1420);border:1px dashed var(--line);
border-radius:13px;padding:20px;text-align:center;margin-top:10px}
.lock .t{font-weight:650} .lock .d{color:var(--muted);font-size:12.5px;margin:6px auto 12px;max-width:520px}
.lock a{display:inline-block;background:var(--accent);color:#05121c;text-decoration:none;
font-weight:650;padding:9px 17px;border-radius:9px;font-size:13px}
footer{margin-top:40px;padding-top:14px;border-top:1px solid var(--line);color:var(--muted);
font-size:11.5px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px}
"""


def _table(headers: List[str], rows: List[List[str]]) -> str:
    head = "".join(f"<th>{html.escape(h)}</th>" for h in headers)
    body = "".join("<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>" for row in rows)
    return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def build(result: Dict, pro: bool = False, version: str = "0.1.0", top: int = 40) -> str:
    deck = result["deck"]
    sev = result["severity"]
    counts = result["counts"]

    cards = [
        ("Score", f"{result['score']:.0f}", "/ 100"),
        ("Findings", str(len(result["findings"])), f"{sev.get('high', 0)} high"),
        ("Slides", str(result["slides"]), f"{result['layouts']} layouts"),
        ("Overflow", str(counts.get("text_overflow", 0)), "text out of box"),
        ("Off-slide", str(counts.get("off_slide", 0)), "outside canvas"),
        ("Notes", str(result["reads"]), "slides with notes"),
    ]
    card_html = "".join(
        f'<div class="card"><div class="l">{l}</div><div class="v">{v}</div>'
        f'<div class="f">{f}</div></div>' for l, v, f in cards
    )

    parts = [f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SlideLint — {html.escape(deck['name'])}</title><style>{CSS}</style></head><body>
<div class="wrap">
<h1>SlideLint — deck review</h1>
<div class="sub">Layout, typography and readability checks. No rendering, no screenshots.</div>
<div class="hero"><div><div class="score" style="color:{result['verdict_color']}">{result['score']:.0f}</div>
<div class="sub">/ 100</div></div>
<div><div class="vt" style="color:{result['verdict_color']}">{html.escape(result['verdict'])}</div>
<div class="vd">{len(result['findings'])} findings · {sev.get('high', 0)} high ·
{sev.get('medium', 0)} medium · {sev.get('low', 0)} low</div></div></div>
<div class="meta">
<span>file <b>{html.escape(deck['name'])}</b></span>
<span>canvas <b>{deck['slide_width']:.2f} × {deck['slide_height']:.2f} in</b></span>
<span>generated <b>{datetime.now():%Y-%m-%d %H:%M}</b></span></div>
<div class="cards">{card_html}</div>
"""]

    rows = summary_rows(result)
    if rows:
        parts.append("<h2>What was found</h2>" + _table(
            ["Check", "Count", "Severity"],
            [[html.escape(r["check"]), str(r["count"]),
              f'<span class="badge {r["severity"]}">{r["severity"]}</span>'] for r in rows]))

    if pro:
        detail = result["findings"][:top]
        parts.append("<h2>Findings</h2>" + _table(
            ["Severity", "Check", "Slide", "Shape", "Detail"],
            [[f'<span class="badge {f["severity"]}">{f["severity"]}</span>',
              html.escape(f["check"]), str(f["slide"]),
              f'<span class="mono">{html.escape(f["target"])}</span>',
              html.escape(f["detail"])] for f in detail]))
    else:
        parts.append("""<h2>Findings</h2><div class="lock"><div class="t">Slide-level findings are Pro</div>
<div class="d">Pro lists every flagged slide with the shape name and the reason — the list you can
hand to whoever owns the deck, plus CSV export for a review tracker.</div>
<a href="https://hymz.gumroad.com/l/slidelint">Get SlideLint Pro</a></div>""")

    parts.append(f"""<footer><span>SlideLint {version} · built on python-pptx (MIT)</span>
<span>Static inspection of the .pptx XML; nothing is rendered.</span></footer></div></body></html>""")
    return "".join(parts)
