"""SlideLint behaviour tests — the checks must fire on a bad deck and stay quiet on a clean one."""

from __future__ import annotations

from pathlib import Path

import pytest

from slidelint import audit, cli, report
from slidelint.licensing import PRODUCT


def _demo(tmp_path: Path) -> Path:
    return cli.demo_deck(tmp_path / "demo_deck.pptx")


def test_product_name_is_the_licence_scope():
    assert PRODUCT == "slidelint"


def test_demo_deck_is_flagged(tmp_path):
    result = audit.audit(_demo(tmp_path))
    assert result["slides"] == 6
    assert result["score"] < 40, "a deck this broken must not score as shippable"
    names = {f["check"] for f in result["findings"]}
    for expected in {"placeholder_text", "text_overflow", "off_slide", "overlapping_text",
                     "missing_title", "tiny_text", "missing_alt_text"}:
        assert expected in names, f"{expected} did not fire"


def test_placeholder_copy_is_high_severity(tmp_path):
    result = audit.audit(_demo(tmp_path))
    hits = [f for f in result["findings"] if f["check"] == "placeholder_text"]
    assert hits and all(f["severity"] == "high" for f in hits)
    assert "Lorem" in hits[0]["detail"] or "lorem" in hits[0]["detail"].lower()


def test_off_slide_reports_the_overflow_amount(tmp_path):
    result = audit.audit(_demo(tmp_path))
    hit = next(f for f in result["findings"] if f["check"] == "off_slide")
    assert "right" in hit["detail"] or "bottom" in hit["detail"]


def test_findings_are_deduplicated(tmp_path):
    result = audit.audit(_demo(tmp_path))
    keys = {(f["check"], f["slide"], f["target"]) for f in result["findings"]}
    assert len(keys) == len(result["findings"])


def test_clean_deck_scores_high(tmp_path):
    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "One idea per slide"
    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(8.0), Inches(1.2))
    para = box.text_frame.paragraphs[0]
    run = para.add_run()
    run.text = "Charts carry the detail; the slide carries the point."
    run.font.size = Pt(18)
    run.font.name = "Calibri"
    slide.notes_slide.notes_text_frame.text = "Say the sentence, then stop talking."
    path = tmp_path / "clean.pptx"
    prs.save(str(path))

    result = audit.audit(path)
    assert result["score"] >= 85
    assert result["verdict"] == "Ship it"


def test_rejects_non_pptx(tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"not a deck")
    with pytest.raises(ValueError):
        audit.audit(path)


def test_report_contains_score_and_lock(tmp_path):
    result = audit.audit(_demo(tmp_path))
    free = report.build(result, pro=False)
    assert "SlideLint — deck review" in free
    assert "Slide-level findings are Pro" in free
    paid = report.build(result, pro=True)
    assert "text_overflow" in paid
    assert "Slide-level findings are Pro" not in paid


def test_summary_rows_only_lists_checks_that_fired(tmp_path):
    result = audit.audit(_demo(tmp_path))
    rows = audit.summary_rows(result)
    assert rows and all(r["count"] > 0 for r in rows)
    assert sum(r["count"] for r in rows) == len(result["findings"])
