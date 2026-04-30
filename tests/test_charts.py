"""Tests for powerpoint_toolkit.charts."""

import pytest

from powerpoint_toolkit.presentation import Presentation
from powerpoint_toolkit.charts import ChartBuilder, CHART_TYPES


def _slide():
    prs = Presentation()
    return prs, prs.add_slide()


# ---------------------------------------------------------------------------
# ChartBuilder construction
# ---------------------------------------------------------------------------

def test_chart_builder_default_type():
    b = ChartBuilder()
    assert b._chart_type is not None


def test_chart_builder_invalid_type():
    with pytest.raises(ValueError, match="Unknown chart type"):
        ChartBuilder("invalid_chart_type")


def test_chart_types_all_valid():
    """All CHART_TYPES keys should construct without error."""
    for chart_type in CHART_TYPES:
        ChartBuilder(chart_type)


# ---------------------------------------------------------------------------
# Fluent API
# ---------------------------------------------------------------------------

def test_fluent_chain_returns_self():
    b = ChartBuilder("column")
    assert b.categories(["A"]) is b
    assert b.add_series("S", [1]) is b
    assert b.title("T") is b
    assert b.legend(True) is b
    assert b.data_labels(False) is b
    assert b.position(1, 1, 6, 4) is b


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def test_build_column_chart():
    prs, slide = _slide()
    chart = (
        ChartBuilder("column")
        .categories(["Q1", "Q2", "Q3", "Q4"])
        .add_series("Revenue", [10, 20, 15, 30])
        .build(slide)
    )
    assert chart is not None


def test_build_bar_chart():
    prs, slide = _slide()
    chart = (
        ChartBuilder("bar")
        .categories(["A", "B", "C"])
        .add_series("S1", [1, 2, 3])
        .add_series("S2", [4, 5, 6])
        .build(slide)
    )
    assert chart is not None


def test_build_pie_chart():
    prs, slide = _slide()
    chart = (
        ChartBuilder("pie")
        .categories(["X", "Y", "Z"])
        .add_series("Share", [30, 40, 30])
        .build(slide)
    )
    assert chart is not None


def test_build_line_chart():
    prs, slide = _slide()
    ChartBuilder("line").categories(["M1"]).add_series("T", [5]).build(slide)


def test_build_chart_with_title():
    prs, slide = _slide()
    chart = (
        ChartBuilder("column")
        .categories(["A"])
        .add_series("S", [1])
        .title("My Title")
        .build(slide)
    )
    assert chart.has_title is True
    assert chart.chart_title.text_frame.text == "My Title"


def test_build_chart_no_legend():
    prs, slide = _slide()
    chart = (
        ChartBuilder("column")
        .categories(["A"])
        .add_series("S", [1])
        .legend(False)
        .build(slide)
    )
    assert chart.has_legend is False


def test_build_chart_with_legend():
    prs, slide = _slide()
    chart = (
        ChartBuilder("column")
        .categories(["A"])
        .add_series("S", [1])
        .legend(True)
        .build(slide)
    )
    assert chart.has_legend is True


def test_build_chart_data_labels():
    prs, slide = _slide()
    chart = (
        ChartBuilder("column")
        .categories(["A"])
        .add_series("S", [1])
        .data_labels(True)
        .build(slide)
    )
    for plot in chart.plots:
        assert plot.has_data_labels is True


def test_build_multiple_charts_on_same_slide():
    prs, slide = _slide()
    for chart_type in ["column", "bar", "line"]:
        ChartBuilder(chart_type).categories(["X"]).add_series("S", [1]).build(slide)
    assert len(slide.shapes) == 3


# ---------------------------------------------------------------------------
# from_dict
# ---------------------------------------------------------------------------

def test_from_dict_basic():
    spec = {
        "type": "column",
        "categories": ["Jan", "Feb"],
        "series": [{"name": "Revenue", "values": [100, 200]}],
        "title": "Monthly",
    }
    b = ChartBuilder.from_dict(spec)
    prs, slide = _slide()
    chart = b.build(slide)
    assert chart.has_title is True


def test_from_dict_defaults():
    b = ChartBuilder.from_dict({})
    assert b is not None


def test_from_dict_legend_and_data_labels():
    b = ChartBuilder.from_dict({"legend": False, "data_labels": True})
    assert b._has_legend is False
    assert b._has_data_labels is True


def test_from_dict_position():
    b = ChartBuilder.from_dict({"position": {"left": 2, "top": 3, "width": 5, "height": 4}})
    assert b._left == 2
    assert b._top == 3
    assert b._width == 5
    assert b._height == 4


def test_from_dict_partial_position():
    b = ChartBuilder.from_dict({"position": {"left": 0.5}})
    assert b._left == 0.5
    assert b._top == 1.5  # default
