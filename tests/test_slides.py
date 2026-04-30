"""Tests for powerpoint_toolkit.slides."""

import pytest
from pathlib import Path

from powerpoint_toolkit.presentation import Presentation
from powerpoint_toolkit.slides import Slide


def _make_slide(layout="blank") -> tuple:
    """Return (prs, slide) with one slide added."""
    prs = Presentation()
    slide = prs.add_slide(layout=layout)
    return prs, slide


# ---------------------------------------------------------------------------
# Slide basic
# ---------------------------------------------------------------------------

def test_slide_repr():
    _, slide = _make_slide()
    assert "Slide" in repr(slide)


def test_slide_has_pptx_slide_property():
    _, slide = _make_slide()
    assert slide.pptx_slide is not None


def test_slide_dimensions():
    prs, slide = _make_slide()
    assert abs(slide.slide_width - 13.33) < 0.01
    assert abs(slide.slide_height - 7.5) < 0.01


# ---------------------------------------------------------------------------
# Background
# ---------------------------------------------------------------------------

def test_set_background_color_returns_self():
    _, slide = _make_slide()
    result = slide.set_background_color("#1F3864")
    assert result is slide


def test_set_background_color_hex_without_hash():
    _, slide = _make_slide()
    slide.set_background_color("FF0000")  # should not raise


# ---------------------------------------------------------------------------
# Title / subtitle
# ---------------------------------------------------------------------------

def test_set_title_on_title_slide():
    _, slide = _make_slide(layout="title")
    slide.set_title("My Title")
    # Find the title placeholder and check text
    found = False
    for ph in slide.pptx_slide.placeholders:
        if ph.placeholder_format.idx == 0:
            assert "My Title" in ph.text_frame.text
            found = True
    assert found


def test_set_title_on_blank_slide_creates_textbox():
    _, slide = _make_slide(layout="blank")
    initial_shapes = len(slide.shapes)
    slide.set_title("Fallback Title")
    assert len(slide.shapes) > initial_shapes


def test_set_subtitle_on_title_slide():
    _, slide = _make_slide(layout="title")
    slide.set_subtitle("My Subtitle")
    # Should not raise; subtitle placeholder exists on title layout


def test_set_title_with_formatting():
    _, slide = _make_slide(layout="title")
    slide.set_title("Bold Blue Title", bold=True, color="#0000FF", font_size=40)


# ---------------------------------------------------------------------------
# Text box
# ---------------------------------------------------------------------------

def test_add_text_box_creates_shape():
    _, slide = _make_slide()
    before = len(slide.shapes)
    slide.add_text_box("Hello, World!")
    assert len(slide.shapes) == before + 1


def test_add_text_box_returns_shape():
    _, slide = _make_slide()
    shape = slide.add_text_box("Hello")
    assert shape is not None


def test_add_text_box_with_all_options():
    _, slide = _make_slide()
    slide.add_text_box(
        "Styled text",
        left=0.5,
        top=1.0,
        width=6.0,
        height=1.5,
        font_size=14,
        font_name="Calibri",
        bold=True,
        italic=True,
        underline=True,
        color="#FF0000",
        alignment="center",
    )


def test_add_text_box_word_wrap():
    _, slide = _make_slide()
    shape = slide.add_text_box("Long text", word_wrap=True)
    assert shape.text_frame.word_wrap is True


# ---------------------------------------------------------------------------
# Bullet list
# ---------------------------------------------------------------------------

def test_add_bullet_list_creates_shape():
    _, slide = _make_slide()
    before = len(slide.shapes)
    slide.add_bullet_list(["Item 1", "Item 2", "Item 3"])
    assert len(slide.shapes) == before + 1


def test_add_bullet_list_correct_paragraph_count():
    _, slide = _make_slide()
    shape = slide.add_bullet_list(["A", "B", "C"])
    assert len(shape.text_frame.paragraphs) == 3


def test_add_bullet_list_empty():
    _, slide = _make_slide()
    shape = slide.add_bullet_list([])
    # Should create a text box with one (empty) paragraph
    assert shape is not None


# ---------------------------------------------------------------------------
# Image
# ---------------------------------------------------------------------------

def test_add_image_file_not_found():
    _, slide = _make_slide()
    with pytest.raises(FileNotFoundError):
        slide.add_image("/nonexistent/path/image.png")


def test_add_image_creates_shape(tmp_path):
    """Create a tiny PNG and add it to a slide."""
    from PIL import Image as PILImage

    img_path = tmp_path / "test.png"
    img = PILImage.new("RGB", (100, 100), color=(255, 0, 0))
    img.save(img_path)

    _, slide = _make_slide()
    before = len(slide.shapes)
    slide.add_image(str(img_path), left=1.0, top=1.0, width=2.0)
    assert len(slide.shapes) == before + 1


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------

def test_add_table_creates_shape():
    _, slide = _make_slide()
    before = len(slide.shapes)
    slide.add_table([["A", "B"], ["1", "2"]])
    assert len(slide.shapes) == before + 1


def test_add_table_cell_text():
    _, slide = _make_slide()
    data = [["Name", "Score"], ["Alice", "95"], ["Bob", "87"]]
    shape = slide.add_table(data)
    table = shape.table
    assert table.cell(0, 0).text == "Name"
    assert table.cell(1, 0).text == "Alice"
    assert table.cell(2, 1).text == "87"


def test_add_table_with_header_styling():
    _, slide = _make_slide()
    slide.add_table(
        [["Name", "Score"], ["Alice", "95"]],
        has_header=True,
        header_bg_color="#1F3864",
        header_font_color="#FFFFFF",
    )


def test_add_table_no_header():
    _, slide = _make_slide()
    data = [["A", "B"], ["C", "D"]]
    shape = slide.add_table(data, has_header=False)
    table = shape.table
    # No bold styling on first row
    run = table.cell(0, 0).text_frame.paragraphs[0].runs[0]
    assert run.font.bold is None or run.font.bold is False


def test_add_table_alternate_row_color():
    _, slide = _make_slide()
    slide.add_table(
        [["H1", "H2"], ["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]],
        alternate_row_color="#DDDDDD",
    )


def test_add_table_none_values():
    _, slide = _make_slide()
    shape = slide.add_table([[None, "B"], ["C", None]])
    table = shape.table
    assert table.cell(0, 0).text == ""
    assert table.cell(1, 1).text == ""


# ---------------------------------------------------------------------------
# Chart (via slide.add_chart)
# ---------------------------------------------------------------------------

def test_add_chart_creates_shape():
    _, slide = _make_slide()
    before = len(slide.shapes)
    slide.add_chart(
        "column",
        ["Q1", "Q2", "Q3"],
        [("Revenue", [10, 20, 15])],
    )
    assert len(slide.shapes) == before + 1


def test_add_chart_with_title():
    _, slide = _make_slide()
    slide.add_chart(
        "bar",
        ["A", "B"],
        [("Series1", [1, 2])],
        title="My Chart",
    )


def test_add_chart_pie():
    _, slide = _make_slide()
    slide.add_chart(
        "pie",
        ["Cats", "Dogs", "Birds"],
        [("Pets", [30, 50, 20])],
    )


def test_add_chart_no_legend():
    _, slide = _make_slide()
    chart = slide.add_chart(
        "line",
        ["Jan", "Feb"],
        [("Trend", [5, 8])],
        has_legend=False,
    )
    assert chart.has_legend is False


def test_add_chart_data_labels():
    _, slide = _make_slide()
    slide.add_chart(
        "column",
        ["X"],
        [("S", [42])],
        has_data_labels=True,
    )


# ---------------------------------------------------------------------------
# Rectangle
# ---------------------------------------------------------------------------

def test_add_rectangle_creates_shape():
    _, slide = _make_slide()
    before = len(slide.shapes)
    slide.add_rectangle()
    assert len(slide.shapes) == before + 1


def test_add_rectangle_with_fill_and_line():
    _, slide = _make_slide()
    slide.add_rectangle(fill_color="#FF0000", line_color="#0000FF")


# ---------------------------------------------------------------------------
# Round-trip
# ---------------------------------------------------------------------------

def test_slide_round_trip(tmp_path):
    prs = Presentation()
    slide = prs.add_slide()
    slide.add_text_box("Round-trip text", left=1, top=1, width=5, height=1)
    out = tmp_path / "rt.pptx"
    prs.save(out)

    prs2 = Presentation(out)
    assert prs2.slide_count == 1
