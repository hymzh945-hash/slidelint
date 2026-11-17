"""`slidelint audit | demo | license`."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from . import audit as audit_mod, licensing, report
from .version import __version__

URL = "https://hymz.gumroad.com/l/slidelint"



def _require_pro(feature: str) -> None:
    print(f"{feature} is part of SlideLint Pro — {URL}")
    raise SystemExit(1)


def demo_deck(path: Path) -> Path:
    """A deck with the mistakes people actually send to clients."""
    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    title_layout = prs.slide_layouts[0]

    # 1 — clean cover
    cover = prs.slides.add_slide(title_layout)
    cover.shapes.title.text = "Q4 Board Review"
    cover.placeholders[1].text = "Finance · November 2026"

    # 2 — placeholder copy left in, one tiny label, one overflowing body box
    s2 = prs.slides.add_slide(blank)
    t2 = s2.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(11), Inches(0.9))
    t2.text_frame.text = "Revenue bridge"
    body = s2.shapes.add_textbox(Inches(0.7), Inches(1.7), Inches(4.2), Inches(1.3))
    body.text_frame.word_wrap = True
    body.text_frame.text = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor "
        "incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud "
        "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure "
        "dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
    )
    tiny = s2.shapes.add_textbox(Inches(0.7), Inches(3.4), Inches(3.0), Inches(0.4))
    tiny.text_frame.text = "source: internal flash report"
    tiny.text_frame.paragraphs[0].runs[0].font.size = Pt(8)

    # 3 — shapes off the canvas, overlapping text, no title at all
    s3 = prs.slides.add_slide(blank)
    a = s3.shapes.add_textbox(Inches(11.6), Inches(6.9), Inches(3.0), Inches(1.2))
    a.text_frame.text = "This box runs off the right edge of the slide"
    b = s3.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(6.0), Inches(1.4))
    b.text_frame.text = "Segment margin"
    c = s3.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(5.5), Inches(1.2))
    c.text_frame.text = "Segment margin by product line, FY24 versus FY25"

    # 4 — font sprawl, reading load, picture with no alt text
    s4 = prs.slides.add_slide(blank)
    head = s4.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(10), Inches(0.8))
    head.text_frame.text = "What changed and why it matters"
    sizes = [Pct for Pct in (9, 10.5, 12, 14, 17, 21, 26, 30)]
    for idx, size in enumerate(sizes):
        box = s4.shapes.add_textbox(Inches(0.7), Inches(1.4 + idx * 0.55), Inches(6), Inches(0.5))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = f"Line {idx + 1} at {size:g}pt"
        run.font.size = Pt(size)
        run.font.name = ["Calibri", "Arial", "Georgia", "Verdana", "Courier New", "Tahoma"][idx % 6]
    para = s4.shapes.add_textbox(Inches(7.4), Inches(1.4), Inches(5.2), Inches(4.6))
    para.text_frame.word_wrap = True
    para.text_frame.text = " ".join(
        ["Pricing moved", "in two regions", "while volume", "held flat"] * 12)
    # a picture with no alt text, straight from an in-memory PNG
    import io
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (160, 90), (18, 26, 41)).save(buf, format="PNG")
    buf.seek(0)
    s4.shapes.add_picture(buf, Inches(7.4), Inches(5.4), Inches(2.2))
    prs.slides.add_slide(blank)  # empty slide, nothing on it

    # 5 — duplicate of slide 2's title, sentence without punctuation
    s5 = prs.slides.add_slide(blank)
    t5 = s5.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(11), Inches(0.9))
    t5.text_frame.text = "Revenue bridge"
    tail = s5.shapes.add_textbox(Inches(0.7), Inches(2.0), Inches(8.0), Inches(1.0))
    tail.text_frame.text = "Next quarter we expect the same shape as the previous three"

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def cmd_audit(args: argparse.Namespace) -> int:
    pro = licensing.is_pro()
    result = audit_mod.audit(args.input)

    print(f"\nSlideLint {__version__} · {result['deck']['name']}")
    print(f"  score      {result['score']:.0f}/100  {result['verdict']}")
    print(f"  findings   {len(result['findings'])} "
          f"({result['severity'].get('high', 0)} high, "
          f"{result['severity'].get('medium', 0)} medium, "
          f"{result['severity'].get('low', 0)} low)")
    print(f"  slides     {result['slides']} · layouts {result['layouts']}")

    for row in audit_mod.summary_rows(result):
        print(f"  - {row['check']:<24} {row['count']:>3}  [{row['severity']}]")

    if pro and args.csv:
        out = Path(args.csv)
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=["severity", "check", "slide", "target", "detail"])
            writer.writeheader()
            for f in result["findings"]:
                writer.writerow({k: f[k] for k in writer.fieldnames})
        print(f"\nfindings -> {out}")

    if args.report:
        if not pro:
            _require_pro("HTML report")
        out = Path(args.report)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(report.build(result, pro=pro, version=__version__, top=args.top), "utf-8")
        print(f"report -> {out}")
    return 0


def cmd_demo(args: argparse.Namespace) -> int:
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    deck = demo_deck(out / "demo_deck.pptx")
    result = audit_mod.audit(deck)

    print(f"{len(result['findings'])} findings · score {result['score']:.0f} · {result['verdict']}")
    for row in audit_mod.summary_rows(result):
        print(f"  - {row['check']}: {row['count']}")

    html = report.build(result, pro=licensing.is_pro(), version=__version__)
    (out / "slidelint-report.html").write_text(html, "utf-8")
    print(f"report -> {out / 'slidelint-report.html'} "
          f"({'Pro' if licensing.is_pro() else 'free tier'})")
    return 0


def cmd_license(args: argparse.Namespace) -> int:
    st = licensing.status()
    if st["valid"]:
        print(f"Pro license · {st['email']} · expires {st['expires'] or 'lifetime'}")
        print(f"source: {licensing.license_path()}")
    else:
        print(f"Free tier · {st['info']}")
        print(f"Pro: {URL}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="slidelint",
                                description="Find what breaks a deck before the room does.")
    p.add_argument("--version", action="version", version=f"slidelint {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    a = sub.add_parser("audit", help="audit a .pptx deck")
    a.add_argument("-i", "--input", required=True)
    a.add_argument("--report", help="write an HTML report (Pro)")
    a.add_argument("--csv", help="write findings to CSV (Pro)")
    a.add_argument("--top", type=int, default=40)
    a.set_defaults(func=cmd_audit)

    d = sub.add_parser("demo", help="build a broken demo deck and audit it")
    d.add_argument("-o", "--out", default="slidelint-demo")
    d.set_defaults(func=cmd_demo)

    l = sub.add_parser("license", help="license status")
    l.set_defaults(func=cmd_license)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
