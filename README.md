# powerpoint-automation-toolkit

A Python toolkit for creating and manipulating PowerPoint (`.pptx`) presentations programmatically, built on top of [python-pptx](https://python-pptx.readthedocs.io/).

## Features

- **Create** brand-new presentations or **open** existing `.pptx` files
- **Add slides** with a variety of built-in layouts
- **Add text boxes** with full font & paragraph formatting
- **Add bullet lists**
- **Add images** (PNG, JPEG, GIF, …)
- **Add tables** with header colouring and alternating row colours
- **Add charts** (column, bar, line, pie, doughnut, area, scatter, …)
- **Add rectangles** and other simple shapes
- **Slide management**: delete, duplicate, reorder
- **Command-line interface** — use the toolkit from the terminal without writing Python
- Fully **tested** with pytest (113 tests)

---

## Installation

```bash
pip install python-pptx Pillow          # runtime dependencies
# or
pip install -r requirements.txt
```

For development (includes pytest and coverage):

```bash
pip install -r requirements-dev.txt
pip install -e .
```

---

## Quick Start (Python API)

```python
from powerpoint_toolkit import Presentation

# Create a new presentation
prs = Presentation()

# Add a title slide
slide = prs.add_slide(layout="title")
slide.set_title("Hello, World!", bold=True, color="#1F3864")
slide.set_subtitle("Built with powerpoint-automation-toolkit")

# Add a blank slide with a text box
slide2 = prs.add_slide(layout="blank")
slide2.set_background_color("#F0F4FF")
slide2.add_text_box(
    "This is a text box",
    left=1.0, top=2.0, width=8.0, height=1.0,
    font_size=18, font_name="Calibri", alignment="center",
)

# Add a slide with a table
slide3 = prs.add_slide(layout="blank")
slide3.add_table(
    [
        ["Name",  "Department", "Score"],
        ["Alice", "Engineering", 95],
        ["Bob",   "Marketing",   87],
        ["Carol", "Design",      92],
    ],
    has_header=True,
    header_bg_color="#1F3864",
    header_font_color="#FFFFFF",
    alternate_row_color="#EBF0FA",
    left=0.5, top=1.5, width=12.0, height=3.0,
)

# Add a slide with a chart
slide4 = prs.add_slide(layout="blank")
slide4.add_chart(
    "column",
    categories=["Q1", "Q2", "Q3", "Q4"],
    series=[
        ("Revenue", [120, 150, 130, 180]),
        ("Profit",  [30,  45,  35,  60]),
    ],
    title="Quarterly Performance",
    left=1.0, top=1.0, width=11.0, height=5.5,
)

# Save the presentation
prs.save("output.pptx")
print(f"Saved {prs.slide_count} slides")
```

---

## Python API Reference

### `Presentation`

```python
from powerpoint_toolkit import Presentation

prs = Presentation()                        # new blank presentation
prs = Presentation("existing.pptx")        # open existing file
prs = Presentation(width_inches=10, height_inches=7.5)
```

| Method / Property | Description |
|---|---|
| `prs.add_slide(layout="blank")` | Add a slide and return a `Slide` object |
| `prs.get_slide(index)` | Return the slide at the given zero-based index |
| `prs.delete_slide(index)` | Delete the slide at the given index |
| `prs.duplicate_slide(index)` | Duplicate a slide and append the copy |
| `prs.reorder_slide(from_index, to_index)` | Move a slide |
| `prs.save(path)` | Save to a `.pptx` file |
| `prs.slide_count` | Number of slides |
| `prs.slides` | List of `Slide` objects |
| `prs.info()` | Dict with metadata |

**Layout names** accepted by `add_slide()`:
`blank`, `title`, `title_and_content`, `title_only`, `two_content`, `comparison`, `content_with_caption`, `picture_with_caption`, `section_header`.

---

### `Slide`

Obtained from `prs.add_slide()` or `prs.get_slide()`.

#### Background

```python
slide.set_background_color("#1F3864")
```

#### Title & Subtitle

```python
slide.set_title("My Title", bold=True, font_size=40, color="#FFFFFF")
slide.set_subtitle("My subtitle text")
```

#### Text Box

```python
slide.add_text_box(
    "Your text here",
    left=1.0, top=1.5, width=8.0, height=1.0,
    font_size=18,
    font_name="Calibri",
    bold=False, italic=False, underline=False,
    color="#333333",
    alignment="left",   # left | center | right | justify
    word_wrap=True,
)
```

#### Bullet List

```python
slide.add_bullet_list(
    ["First point", "Second point", "Third point"],
    left=1.0, top=2.0, width=8.0, height=4.0,
    font_size=16,
)
```

#### Image

```python
slide.add_image(
    "path/to/image.png",
    left=1.0, top=1.5,
    width=4.0,          # set width; height scales proportionally
)
```

#### Table

```python
slide.add_table(
    data=[
        ["Name", "Score"],   # header row
        ["Alice", 95],
        ["Bob",   87],
    ],
    has_header=True,
    header_bg_color="#1F3864",
    header_font_color="#FFFFFF",
    alternate_row_color="#EBF0FA",
    font_size=14,
    left=0.5, top=1.5, width=12.0, height=3.0,
)
```

#### Chart

```python
slide.add_chart(
    chart_type="column",        # see table below
    categories=["Q1", "Q2", "Q3"],
    series=[
        ("Revenue", [100, 200, 150]),
        ("Profit",  [30,  60,  45]),
    ],
    title="Quarterly Numbers",
    has_legend=True,
    has_data_labels=False,
    left=1.0, top=1.5, width=10.0, height=5.0,
)
```

**Chart types**: `column`, `column_stacked`, `bar`, `bar_stacked`, `line`, `line_markers`, `pie`, `doughnut`, `area`, `scatter`.

#### Rectangle

```python
slide.add_rectangle(
    left=1.0, top=1.0, width=4.0, height=2.0,
    fill_color="#1F3864",
    line_color="#FFFFFF",
)
```

---

### `ChartBuilder` (fluent API)

For more complex chart scenarios, use the `ChartBuilder` directly:

```python
from powerpoint_toolkit.charts import ChartBuilder

chart = (
    ChartBuilder("column")
    .categories(["Q1", "Q2", "Q3", "Q4"])
    .add_series("Revenue", [120, 150, 130, 180])
    .add_series("Profit",  [30,  45,  35,  60])
    .title("Quarterly Performance")
    .legend(True)
    .data_labels(False)
    .position(left=1.0, top=1.5, width=10.0, height=5.0)
    .build(slide)
)
```

Or build from a dict:

```python
chart = ChartBuilder.from_dict({
    "type": "bar",
    "categories": ["A", "B", "C"],
    "series": [{"name": "Values", "values": [10, 20, 30]}],
    "title": "My Bar Chart",
    "legend": True,
    "position": {"left": 1, "top": 2, "width": 8, "height": 4},
}).build(slide)
```

---

## Command-Line Interface

The toolkit ships with a CLI. After installation run:

```bash
pptx-toolkit --help
# or
python -m powerpoint_toolkit --help
```

### `create` — create a new presentation

```bash
pptx-toolkit create -o slides.pptx
pptx-toolkit create -o slides.pptx --width 10 --height 7.5
```

### `add-slide` — append a slide

```bash
pptx-toolkit add-slide -f slides.pptx --layout title \
    --title "Hello World" --subtitle "My first slide"

pptx-toolkit add-slide -f slides.pptx --layout blank \
    --background "#1F3864"
```

### `add-text` — add a text box

```bash
pptx-toolkit add-text -f slides.pptx --slide 0 \
    --text "Important message" \
    --font-size 24 --bold --color "#FF0000" \
    --alignment center
```

### `add-image` — add an image

```bash
pptx-toolkit add-image -f slides.pptx --slide 0 \
    --image photo.png --left 1 --top 2 --width 6
```

### `add-table` — add a table

```bash
pptx-toolkit add-table -f slides.pptx --slide 0 \
    --rows "Name,Department,Score" \
    --rows "Alice,Engineering,95" \
    --rows "Bob,Marketing,87" \
    --header-bg "#1F3864" --header-color "#FFFFFF"
```

### `add-chart` — add a chart

```bash
pptx-toolkit add-chart -f slides.pptx --slide 0 \
    --type column \
    --categories "Q1,Q2,Q3,Q4" \
    --series "Revenue:120,150,130,180" \
    --series "Profit:30,45,35,60" \
    --title "Quarterly Performance"
```

### `info` — show presentation metadata

```bash
pptx-toolkit info -f slides.pptx
pptx-toolkit info -f slides.pptx --json
```

---

## Running Tests

```bash
pip install -r requirements-dev.txt
pytest
```

---

## Project Structure

```
powerpoint_toolkit/
├── __init__.py        # Public API (Presentation, Slide, ChartBuilder)
├── __main__.py        # python -m powerpoint_toolkit entry point
├── presentation.py    # Presentation class
├── slides.py          # Slide class with content helpers
├── shapes.py          # Text / paragraph formatting utilities
├── charts.py          # ChartBuilder
└── cli.py             # Command-line interface

tests/
├── test_presentation.py
├── test_slides.py
├── test_charts.py
├── test_shapes.py
└── test_cli.py
```

---

## License

MIT
