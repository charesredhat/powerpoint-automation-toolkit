"""Tests for powerpoint_toolkit.presentation."""

import pytest
from pathlib import Path

from powerpoint_toolkit.presentation import Presentation


# ---------------------------------------------------------------------------
# Creation
# ---------------------------------------------------------------------------

def test_new_presentation_has_zero_slides():
    prs = Presentation()
    assert prs.slide_count == 0


def test_new_presentation_default_size():
    prs = Presentation()
    assert abs(prs.slide_width - 13.33) < 0.01
    assert abs(prs.slide_height - 7.5) < 0.01


def test_new_presentation_custom_size():
    prs = Presentation(width_inches=10.0, height_inches=7.5)
    assert abs(prs.slide_width - 10.0) < 0.01
    assert abs(prs.slide_height - 7.5) < 0.01


# ---------------------------------------------------------------------------
# Slide management
# ---------------------------------------------------------------------------

def test_add_slide_increments_count():
    prs = Presentation()
    prs.add_slide()
    assert prs.slide_count == 1


def test_add_multiple_slides():
    prs = Presentation()
    for _ in range(5):
        prs.add_slide()
    assert prs.slide_count == 5


def test_add_slide_returns_slide_object():
    from powerpoint_toolkit.slides import Slide

    prs = Presentation()
    slide = prs.add_slide()
    assert isinstance(slide, Slide)


def test_add_slide_all_named_layouts():
    layouts = [
        "blank", "title", "title_and_content", "title_only",
        "two_content", "comparison", "content_with_caption",
        "picture_with_caption", "section_header",
    ]
    prs = Presentation()
    for layout in layouts:
        prs.add_slide(layout=layout)
    assert prs.slide_count == len(layouts)


def test_add_slide_integer_layout():
    prs = Presentation()
    prs.add_slide(layout=0)
    assert prs.slide_count == 1


def test_add_slide_invalid_layout_raises():
    prs = Presentation()
    with pytest.raises(ValueError, match="Unknown layout"):
        prs.add_slide(layout="nonexistent_layout")


def test_get_slide_valid():
    prs = Presentation()
    prs.add_slide()
    slide = prs.get_slide(0)
    assert slide is not None


def test_get_slide_out_of_range_raises():
    prs = Presentation()
    with pytest.raises(IndexError):
        prs.get_slide(0)


def test_delete_slide():
    prs = Presentation()
    prs.add_slide()
    prs.add_slide()
    prs.delete_slide(0)
    assert prs.slide_count == 1


def test_delete_slide_out_of_range_raises():
    prs = Presentation()
    with pytest.raises(IndexError):
        prs.delete_slide(0)


def test_duplicate_slide():
    prs = Presentation()
    prs.add_slide()
    prs.duplicate_slide(0)
    assert prs.slide_count == 2


def test_reorder_slide():
    prs = Presentation()
    s0 = prs.add_slide()
    s1 = prs.add_slide()
    # Get underlying pptx ids before reorder
    id_0_before = prs.pptx.slides[0].shapes._spTree
    prs.reorder_slide(0, 1)
    assert prs.slide_count == 2


def test_reorder_slide_out_of_range():
    prs = Presentation()
    prs.add_slide()
    with pytest.raises(IndexError):
        prs.reorder_slide(0, 5)


# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------

def test_save_creates_file(tmp_path):
    out = tmp_path / "test.pptx"
    prs = Presentation()
    prs.add_slide()
    prs.save(out)
    assert out.exists()
    assert out.stat().st_size > 0


def test_save_creates_parent_dirs(tmp_path):
    out = tmp_path / "a" / "b" / "test.pptx"
    prs = Presentation()
    prs.save(out)
    assert out.exists()


def test_save_no_path_raises():
    prs = Presentation()
    with pytest.raises(ValueError, match="No file path"):
        prs.save()


def test_open_saved_file(tmp_path):
    out = tmp_path / "roundtrip.pptx"
    prs = Presentation()
    prs.add_slide()
    prs.add_slide()
    prs.save(out)

    prs2 = Presentation(out)
    assert prs2.slide_count == 2


def test_save_returns_path_object(tmp_path):
    out = tmp_path / "test.pptx"
    prs = Presentation()
    result = prs.save(out)
    assert isinstance(result, Path)
    assert result == out


# ---------------------------------------------------------------------------
# Info
# ---------------------------------------------------------------------------

def test_info_dict(tmp_path):
    out = tmp_path / "info.pptx"
    prs = Presentation()
    prs.add_slide()
    prs.save(out)

    info = prs.info()
    assert info["slide_count"] == 1
    assert "slide_width_inches" in info
    assert "slide_height_inches" in info


def test_repr():
    prs = Presentation()
    r = repr(prs)
    assert "Presentation" in r
    assert "slides=0" in r
