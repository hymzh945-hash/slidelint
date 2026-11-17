"""Read a .pptx deck and report what breaks it in front of an audience."""

from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Dict, List, Tuple

from pptx import Presentation
from pptx.util import Emu

EMU_PER_INCH = 914400
SEVERITY_WEIGHT = {"high": 20, "medium": 8, "low": 2}
CHECK_CAP = {"placeholder_text": 3, "text_overflow": 4, "off_slide": 4, "overlapping_text": 4,
             "font_outlier": 3, "font_size_sprawl": 3, "tiny_text": 4, "missing_title": 4,
             "duplicate_title": 3, "reading_load": 2, "missing_alt_text": 3,
             "no_speaker_notes": 2, "unterminated_sentence": 3}

PLACEHOLDER_TEXT = re.compile(r"click to add|click to edit|lorem ipsum|tbd|todo|xxx", re.IGNORECASE)
SENTENCE_END = (".", "!", "?", ":", "。", "！", "？", "；")
MIN_READABLE_PT = 12.0
MAX_WORDS_PER_SLIDE = 90
TITLE_CASE_WORDS = 6


def _finding(check: str, severity: str, slide: int, target: str, detail: str) -> Dict:
    return {"check": check, "severity": severity, "slide": slide, "target": target,
            "detail": detail, "weight": SEVERITY_WEIGHT[severity]}


def _inches(emu) -> float:
    return float(emu) / EMU_PER_INCH if emu is not None else 0.0


def _shape_text(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(p.text for p in shape.text_frame.paragraphs).strip()


def _runs(shape) -> List:
    out = []
    if not getattr(shape, "has_text_frame", False):
        return out
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.text.strip():
                out.append(run)
    return out


def _alt_text(shape) -> str:
    try:
        for node in shape._element.xpath(".//p:cNvPr"):
            value = node.get("descr") or ""
            if value:
                return value
    except Exception:
        return ""
    return ""


def _overlap(a, b) -> float:
    """Return the share of the smaller box covered by the intersection."""
    ax1, ay1 = _inches(a.left), _inches(a.top)
    ax2, ay2 = ax1 + _inches(a.width), ay1 + _inches(a.height)
    bx1, by1 = _inches(b.left), _inches(b.top)
    bx2, by2 = bx1 + _inches(b.width), by1 + _inches(b.height)
    dx = max(0.0, min(ax2, bx2) - max(ax1, bx1))
    dy = max(0.0, min(ay2, by2) - max(ay1, by1))
    if dx == 0 or dy == 0:
        return 0.0
    inter = dx * dy
    smaller = min(_inches(a.width) * _inches(a.height), _inches(b.width) * _inches(b.height))
    return inter / smaller if smaller else 0.0


def _capacity(shape, size_pt: float) -> int:
    """Rough character capacity for a text box at a given font size."""
    width_in, height_in = _inches(shape.width), _inches(shape.height)
    char_width_px = max(4.0, size_pt * 0.55 * 96 / 72)
    line_height_px = max(10.0, size_pt * 1.2 * 96 / 72)
    chars_per_line = max(4, int((width_in * 96) / char_width_px))
    lines = max(1, int((height_in * 96) / line_height_px))
    return chars_per_line * lines


def _slide_title(shapes: List[Dict]) -> str:
    """A real title placeholder wins; otherwise take the widest text box near the top."""
    for shape in shapes:
        if "TITLE" in shape["placeholder_type"].upper() and shape["text"]:
            return shape["text"].splitlines()[0]
    candidates = [
        s for s in shapes
        if s["is_text"] and s["top"] <= 1.6 and s["width"] >= 4.0
        and len(s["text"]) <= 90 and s["text"].count("\n") <= 1
    ]
    if not candidates:
        return ""
    best = max(candidates, key=lambda s: s["width"] * s["height"])
    return best["text"].splitlines()[0]


def load(path: str | Path) -> Dict:
    path = Path(path)
    if path.suffix.lower() not in {".pptx", ".pptm"}:
        raise ValueError("only .pptx / .pptm decks are supported")
    prs = Presentation(str(path))

    slides = []
    for index, slide in enumerate(prs.slides, start=1):
        shapes = []
        for shape in slide.shapes:
            shapes.append({
                "shape": shape,
                "name": shape.name,
                "text": _shape_text(shape),
                "is_text": getattr(shape, "has_text_frame", False) and bool(_shape_text(shape)),
                "is_picture": shape.shape_type == 13 or shape.name.lower().startswith("picture"),
                "is_placeholder": bool(getattr(shape, "is_placeholder", False)),
                "placeholder_type": (
                    str(shape.placeholder_format.type) if getattr(shape, "is_placeholder", False) else ""
                ),
                "alt": _alt_text(shape),
                "left": _inches(shape.left),
                "top": _inches(shape.top),
                "width": _inches(shape.width),
                "height": _inches(shape.height),
            })
        has_notes = slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip() != ""
        slides.append({
            "number": index,
            "shapes": shapes,
            "notes": has_notes,
            "title": _slide_title(shapes),
        })

    return {
        "path": str(path),
        "name": path.name,
        "slides": slides,
        "slide_width": _inches(prs.slide_width),
        "slide_height": _inches(prs.slide_height),
        "layouts": len(prs.slide_layouts),
    }


def check_placeholders(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        for shape in slide["shapes"]:
            text = shape["text"]
            if not text:
                continue
            if PLACEHOLDER_TEXT.search(text):
                out.append(_finding("placeholder_text", "high", slide["number"], shape["name"],
                                    f'placeholder copy left in slide: "{text[:48]}"'))
    return out


def check_overflow(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        for shape in slide["shapes"]:
            if not shape["is_text"]:
                continue
            sizes = [r.font.size.pt for r in _runs(shape["shape"]) if r.font.size is not None]
            size = max(sizes) if sizes else 18.0
            text = shape["text"]
            cap = _capacity(shape["shape"], size)
            if len(text) > cap * 1.15:
                out.append(_finding(
                    "text_overflow", "high", slide["number"], shape["name"],
                    f'{len(text)} chars in a box that fits about {cap} at {size:.0f}pt'))
    return out


def check_offslide(deck: Dict) -> List[Dict]:
    out = []
    width, height = deck["slide_width"], deck["slide_height"]
    for slide in deck["slides"]:
        for shape in slide["shapes"]:
            over_right = shape["left"] + shape["width"] - width
            over_bottom = shape["top"] + shape["height"] - height
            if shape["left"] < -0.05 or shape["top"] < -0.05 or over_right > 0.05 or over_bottom > 0.05:
                out.append(_finding(
                    "off_slide", "high", slide["number"], shape["name"],
                    f"shape sits outside the {width:.2f}x{height:.2f}in canvas "
                    f"(right +{over_right:.2f}in, bottom +{over_bottom:.2f}in)"))
    return out


def check_overlap(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        texts = [s for s in slide["shapes"] if s["is_text"]]
        for i in range(len(texts)):
            for j in range(i + 1, len(texts)):
                a, b = texts[i], texts[j]
                shared = _overlap(a["shape"], b["shape"])
                if shared >= 0.30:
                    out.append(_finding(
                        "overlapping_text", "medium", slide["number"], f"{a['name']} / {b['name']}",
                        f"two text boxes overlap {shared * 100:.0f}% — one will cover the other"))
    return out


def check_fonts(deck: Dict) -> List[Dict]:
    out = []
    fonts: Dict[str, List[Tuple[int, str]]] = {}
    for slide in deck["slides"]:
        for shape in slide["shapes"]:
            for run in _runs(shape["shape"]):
                name = run.font.name or "(theme)"
                fonts.setdefault(name, []).append((slide["number"], shape["name"]))
    if len(fonts) > 2:
        for name, where in fonts.items():
            if len(where) <= 1:
                out.append(_finding("font_outlier", "medium", where[0][0], where[0][1],
                                    f'"{name}" appears once in the whole deck'))
    return out


def check_font_sizes(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        sizes: Dict[float, List[str]] = {}
        for shape in slide["shapes"]:
            for run in _runs(shape["shape"]):
                if run.font.size is not None and run.text.strip():
                    sizes.setdefault(round(run.font.size.pt, 1), []).append(shape["name"])
        if len(sizes) > 4:
            out.append(_finding("font_size_sprawl", "medium", slide["number"], "whole slide",
                                f"{len(sizes)} distinct font sizes on one slide: "
                                + ", ".join(f"{s:g}pt" for s in sorted(sizes)[:8])))
        for size, where in sizes.items():
            if size < MIN_READABLE_PT:
                out.append(_finding("tiny_text", "medium", slide["number"], where[0],
                                    f"{size:g}pt text will not survive a projector or a phone"))
    return out


def check_titles(deck: Dict) -> List[Dict]:
    out = []
    seen: Dict[str, int] = {}
    for slide in deck["slides"]:
        title = slide["title"]
        if slide["number"] == 1:
            continue
        if not title:
            out.append(_finding("missing_title", "medium", slide["number"], "slide",
                                "no title placeholder text — the audience has no anchor"))
            continue
        key = title.strip().lower()
        if key in seen:
            out.append(_finding("duplicate_title", "low", slide["number"], title[:40],
                                f"same title already used on slide {seen[key]}"))
        else:
            seen[key] = slide["number"]
    return out


def check_reading_load(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        words = sum(len(s["text"].split()) for s in slide["shapes"] if s["is_text"])
        if words > MAX_WORDS_PER_SLIDE:
            out.append(_finding("reading_load", "low", slide["number"], "whole slide",
                                f"{words} words on one slide (over {MAX_WORDS_PER_SLIDE})"))
    return out


def check_image_alt(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        for shape in slide["shapes"]:
            if not shape["is_picture"]:
                continue
            alt = shape["alt"].strip()
            # Office seeds the alt text with the file name; that is not a description.
            if not alt or re.match(r".*\.(png|jpe?g|gif|webp|svg|emf|bmp)$", alt, re.IGNORECASE):
                out.append(_finding("missing_alt_text", "medium", slide["number"], shape["name"],
                                    "image alt text is empty or just the file name"))
    return out


def check_notes(deck: Dict) -> List[Dict]:
    out = []
    for slide in deck["slides"]:
        if slide["number"] == 1:
            continue
        if not slide["notes"]:
            out.append(_finding("no_speaker_notes", "low", slide["number"], "slide",
                                "no speaker notes — nobody can hand this slide to a colleague"))
    return out


def check_punctuation(deck: Dict) -> List[Dict]:
    """One finding per slide: the first long sentence with no closing punctuation."""
    out = []
    for slide in deck["slides"]:
        for shape in slide["shapes"]:
            if not shape["is_text"]:
                continue
            for para in shape["shape"].text_frame.paragraphs:
                text = para.text.strip()
                if len(text.split()) >= 8 and not text.endswith(SENTENCE_END):
                    out.append(_finding("unterminated_sentence", "low", slide["number"], shape["name"],
                                        f'sentence has no closing punctuation: "{text[:52]}"'))
                    break
            else:
                continue
            break
    return out


CHECKS = [
    ("placeholder_text", 20, check_placeholders),
    ("text_overflow", 20, check_overflow),
    ("off_slide", 20, check_offslide),
    ("overlapping_text", 8, check_overlap),
    ("font_outlier", 8, check_fonts),
    ("font_size_sprawl", 8, check_font_sizes),
    ("tiny_text", 8, check_font_sizes),
    ("missing_title", 8, check_titles),
    ("duplicate_title", 2, check_titles),
    ("reading_load", 2, check_reading_load),
    ("missing_alt_text", 8, check_image_alt),
    ("no_speaker_notes", 2, check_notes),
    ("unterminated_sentence", 2, check_punctuation),
]


def audit(path: str | Path) -> Dict:
    deck = load(path)
    findings: List[Dict] = []
    for _name, _weight, fn in CHECKS:
        findings.extend(fn(deck))
    # de-duplicate by (check, slide, target)
    seen, unique = set(), []
    for f in findings:
        key = (f["check"], f["slide"], f["target"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(f)

    penalty = 0.0
    for name, _weight, _fn in CHECKS:
        count = sum(1 for f in unique if f["check"] == name)
        if not count:
            continue
        penalty += min(count, CHECK_CAP.get(name, 3)) * SEVERITY_WEIGHT[
            next(f["severity"] for f in unique if f["check"] == name)]
    # exponential decay keeps a 40-finding deck distinguishable from a 12-finding one
    score = 100.0 * math.exp(-penalty / 120.0)
    if score >= 85:
        verdict, color = "Ship it", "#34d399"
    elif score >= 65:
        verdict, color = "Fixable in an hour", "#fbbf24"
    else:
        verdict, color = "Will embarrass you live", "#f87171"

    severity = {"high": 0, "medium": 0, "low": 0}
    for f in unique:
        severity[f["severity"]] += 1

    return {
        "deck": {"name": deck["name"], "slides": deck["slides"],
                 "slide_width": deck["slide_width"], "slide_height": deck["slide_height"]},
        "findings": unique,
        "score": score,
        "verdict": verdict,
        "verdict_color": color,
        "severity": severity,
        "slides": len(deck["slides"]),
        "layouts": deck["layouts"],
        "counts": {name: sum(1 for f in unique if f["check"] == name) for name, _w, _f in CHECKS},
        "reads": sum(1 for s in deck["slides"] if s["notes"]),
    }


def summary_rows(result: Dict) -> List[Dict]:
    rows = []
    for name, _weight, _fn in CHECKS:
        count = result["counts"].get(name, 0)
        if not count:
            continue
        severity = next((f["severity"] for f in result["findings"] if f["check"] == name), "low")
        rows.append({"check": name.replace("_", " "), "count": count, "severity": severity})
    return rows
