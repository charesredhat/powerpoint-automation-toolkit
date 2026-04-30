"""Tests for the powerpoint_toolkit CLI."""

import json
import pytest
from pathlib import Path

from powerpoint_toolkit.cli import main


def run(*args):
    """Invoke the CLI and return the exit code."""
    return main(list(args))


# ---------------------------------------------------------------------------
# create
# ---------------------------------------------------------------------------

def test_cli_create(tmp_path):
    out = str(tmp_path / "new.pptx")
    assert run("create", "-o", out) == 0
    assert Path(out).exists()


def test_cli_create_custom_size(tmp_path):
    out = str(tmp_path / "custom.pptx")
    assert run("create", "-o", out, "--width", "10", "--height", "7.5") == 0


# ---------------------------------------------------------------------------
# add-slide
# ---------------------------------------------------------------------------

def test_cli_add_slide(tmp_path):
    out = str(tmp_path / "slides.pptx")
    run("create", "-o", out)
    assert run("add-slide", "-f", out, "--layout", "blank") == 0

    from powerpoint_toolkit.presentation import Presentation
    prs = Presentation(out)
    assert prs.slide_count == 1


def test_cli_add_slide_title_layout_with_text(tmp_path):
    out = str(tmp_path / "titled.pptx")
    run("create", "-o", out)
    assert run(
        "add-slide", "-f", out,
        "--layout", "title",
        "--title", "Hello World",
        "--subtitle", "Sub text",
    ) == 0


def test_cli_add_slide_with_background(tmp_path):
    out = str(tmp_path / "bg.pptx")
    run("create", "-o", out)
    assert run("add-slide", "-f", out, "--background", "#1F3864") == 0


def test_cli_add_slide_separate_output(tmp_path):
    src = str(tmp_path / "src.pptx")
    dst = str(tmp_path / "dst.pptx")
    run("create", "-o", src)
    assert run("add-slide", "-f", src, "-o", dst) == 0
    assert Path(dst).exists()


# ---------------------------------------------------------------------------
# add-text
# ---------------------------------------------------------------------------

def test_cli_add_text(tmp_path):
    out = str(tmp_path / "text.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-text", "-f", out, "--slide", "0", "--text", "Hello CLI",
    ) == 0


def test_cli_add_text_with_formatting(tmp_path):
    out = str(tmp_path / "fmt.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-text", "-f", out,
        "--text", "Styled",
        "--font-size", "24",
        "--bold",
        "--color", "#FF0000",
        "--alignment", "center",
    ) == 0


# ---------------------------------------------------------------------------
# add-image
# ---------------------------------------------------------------------------

def test_cli_add_image_missing_file(tmp_path):
    out = str(tmp_path / "img.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    # Non-existent image should fail gracefully (exit code 1)
    code = run("add-image", "-f", out, "--image", "/nonexistent/image.png")
    assert code == 1


def test_cli_add_image_success(tmp_path):
    from PIL import Image as PILImage

    img_path = tmp_path / "test.png"
    PILImage.new("RGB", (100, 100), color=(0, 128, 0)).save(img_path)

    out = str(tmp_path / "img.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run("add-image", "-f", out, "--image", str(img_path), "--width", "3") == 0


# ---------------------------------------------------------------------------
# add-table
# ---------------------------------------------------------------------------

def test_cli_add_table(tmp_path):
    out = str(tmp_path / "tbl.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-table", "-f", out,
        "--rows", "Name,Score",
        "--rows", "Alice,95",
        "--rows", "Bob,87",
    ) == 0


def test_cli_add_table_no_header(tmp_path):
    out = str(tmp_path / "tbl2.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-table", "-f", out,
        "--rows", "A,B",
        "--rows", "C,D",
        "--no-header",
    ) == 0


def test_cli_add_table_with_colors(tmp_path):
    out = str(tmp_path / "tblc.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-table", "-f", out,
        "--rows", "H1,H2",
        "--rows", "D1,D2",
        "--header-bg", "#1F3864",
        "--header-color", "#FFFFFF",
        "--alt-row-color", "#F0F0F0",
    ) == 0


# ---------------------------------------------------------------------------
# add-chart
# ---------------------------------------------------------------------------

def test_cli_add_chart_column(tmp_path):
    out = str(tmp_path / "chart.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-chart", "-f", out,
        "--type", "column",
        "--categories", "Q1,Q2,Q3,Q4",
        "--series", "Revenue:10,20,15,30",
        "--title", "Quarterly",
    ) == 0


def test_cli_add_chart_multiple_series(tmp_path):
    out = str(tmp_path / "chart2.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-chart", "-f", out,
        "--type", "bar",
        "--categories", "A,B",
        "--series", "S1:1,2",
        "--series", "S2:3,4",
    ) == 0


def test_cli_add_chart_pie(tmp_path):
    out = str(tmp_path / "pie.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run(
        "add-chart", "-f", out,
        "--type", "pie",
        "--categories", "X,Y,Z",
        "--series", "Share:30,40,30",
        "--no-legend",
        "--data-labels",
    ) == 0


# ---------------------------------------------------------------------------
# info
# ---------------------------------------------------------------------------

def test_cli_info(tmp_path, capsys):
    out = str(tmp_path / "info.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    assert run("info", "-f", out) == 0
    captured = capsys.readouterr()
    assert "Slides:" in captured.out


def test_cli_info_json(tmp_path, capsys):
    out = str(tmp_path / "info.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    capsys.readouterr()  # clear accumulated output from create/add-slide
    assert run("info", "-f", out, "--json") == 0
    captured = capsys.readouterr()
    data = json.loads(captured.out)
    assert data["slide_count"] == 1


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------

def test_cli_add_text_slide_out_of_range(tmp_path):
    out = str(tmp_path / "err.pptx")
    run("create", "-o", out)
    run("add-slide", "-f", out)
    code = run("add-text", "-f", out, "--slide", "99", "--text", "hi")
    assert code == 1
