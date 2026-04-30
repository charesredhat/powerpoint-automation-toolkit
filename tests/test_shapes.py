"""Tests for powerpoint_toolkit.shapes utilities."""

import pytest

from powerpoint_toolkit.shapes import (
    hex_to_rgb,
    apply_text_format,
    apply_paragraph_format,
    parse_position,
)
from pptx.dml.color import RGBColor


# ---------------------------------------------------------------------------
# hex_to_rgb
# ---------------------------------------------------------------------------

def test_hex_to_rgb_with_hash():
    result = hex_to_rgb("#FF5733")
    assert result == RGBColor(0xFF, 0x57, 0x33)


def test_hex_to_rgb_without_hash():
    result = hex_to_rgb("1F3864")
    assert result == RGBColor(0x1F, 0x38, 0x64)


def test_hex_to_rgb_black():
    assert hex_to_rgb("#000000") == RGBColor(0, 0, 0)


def test_hex_to_rgb_white():
    assert hex_to_rgb("#FFFFFF") == RGBColor(255, 255, 255)


def test_hex_to_rgb_lowercase():
    assert hex_to_rgb("#ff0000") == RGBColor(255, 0, 0)


def test_hex_to_rgb_invalid_length():
    with pytest.raises(ValueError, match="Invalid hex colour"):
        hex_to_rgb("FFF")


def test_hex_to_rgb_invalid_chars():
    with pytest.raises(ValueError):
        hex_to_rgb("#ZZZZZZ")


# ---------------------------------------------------------------------------
# apply_text_format (via a mock run)
# ---------------------------------------------------------------------------

class _MockFont:
    bold = None
    italic = None
    underline = None
    size = None
    name = None

    class _Color:
        rgb = None

    color = _Color()


class _MockRun:
    font = _MockFont()


def _fresh_run():
    """Return a fresh mock run (new instance per call)."""
    r = _MockRun()
    r.font = _MockFont()
    r.font.color = _MockFont._Color()
    return r


def test_apply_text_format_bold():
    run = _fresh_run()
    apply_text_format(run, bold=True)
    assert run.font.bold is True


def test_apply_text_format_italic():
    run = _fresh_run()
    apply_text_format(run, italic=True)
    assert run.font.italic is True


def test_apply_text_format_underline():
    run = _fresh_run()
    apply_text_format(run, underline=True)
    assert run.font.underline is True


def test_apply_text_format_font_size():
    from pptx.util import Pt

    run = _fresh_run()
    apply_text_format(run, font_size=18)
    assert run.font.size == Pt(18)


def test_apply_text_format_font_name():
    run = _fresh_run()
    apply_text_format(run, font_name="Arial")
    assert run.font.name == "Arial"


def test_apply_text_format_color():
    run = _fresh_run()
    apply_text_format(run, color="#FF0000")
    assert run.font.color.rgb == RGBColor(255, 0, 0)


def test_apply_text_format_no_args():
    """Calling with no formatting args should be a no-op."""
    run = _fresh_run()
    apply_text_format(run)
    assert run.font.bold is None


# ---------------------------------------------------------------------------
# apply_paragraph_format (via mock paragraph)
# ---------------------------------------------------------------------------

class _MockParagraphFormat:
    space_before = None
    space_after = None
    line_spacing = None


class _MockParagraph:
    alignment = None
    paragraph_format = _MockParagraphFormat()


def _fresh_para():
    p = _MockParagraph()
    p.paragraph_format = _MockParagraphFormat()
    return p


def test_apply_paragraph_format_alignment_left():
    from pptx.enum.text import PP_ALIGN

    para = _fresh_para()
    apply_paragraph_format(para, alignment="left")
    assert para.alignment == PP_ALIGN.LEFT


def test_apply_paragraph_format_alignment_center():
    from pptx.enum.text import PP_ALIGN

    para = _fresh_para()
    apply_paragraph_format(para, alignment="center")
    assert para.alignment == PP_ALIGN.CENTER


def test_apply_paragraph_format_alignment_right():
    from pptx.enum.text import PP_ALIGN

    para = _fresh_para()
    apply_paragraph_format(para, alignment="right")
    assert para.alignment == PP_ALIGN.RIGHT


def test_apply_paragraph_format_alignment_justify():
    from pptx.enum.text import PP_ALIGN

    para = _fresh_para()
    apply_paragraph_format(para, alignment="justify")
    assert para.alignment == PP_ALIGN.JUSTIFY


def test_apply_paragraph_format_invalid_alignment():
    para = _fresh_para()
    with pytest.raises(ValueError, match="Unknown alignment"):
        apply_paragraph_format(para, alignment="diagonal")


def test_apply_paragraph_format_spacing():
    from pptx.util import Pt

    para = _fresh_para()
    apply_paragraph_format(para, space_before=6, space_after=4, line_spacing=14)
    assert para.paragraph_format.space_before == Pt(6)
    assert para.paragraph_format.space_after == Pt(4)
    assert para.paragraph_format.line_spacing == Pt(14)


# ---------------------------------------------------------------------------
# parse_position
# ---------------------------------------------------------------------------

def test_parse_position_returns_emu():
    from pptx.util import Inches

    left, top, width, height = parse_position(1, 2, 8, 4)
    assert left == Inches(1)
    assert top == Inches(2)
    assert width == Inches(8)
    assert height == Inches(4)
