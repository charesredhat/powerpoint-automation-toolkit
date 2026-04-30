"""Command-line interface for the PowerPoint Automation Toolkit.

Usage examples::

    # Create a new presentation
    python -m powerpoint_toolkit create -o slides.pptx

    # Add a title slide
    python -m powerpoint_toolkit add-slide -f slides.pptx --layout title \\
        --title "Hello World" --subtitle "My first slide"

    # Add a text box
    python -m powerpoint_toolkit add-text -f slides.pptx --slide 0 \\
        --text "Body text here" --font-size 18

    # Add a table (CSV-style rows)
    python -m powerpoint_toolkit add-table -f slides.pptx --slide 0 \\
        --rows "Name,Score" --rows "Alice,95" --rows "Bob,87"

    # Add a bar chart
    python -m powerpoint_toolkit add-chart -f slides.pptx --slide 0 \\
        --type column --categories "Q1,Q2,Q3,Q4" \\
        --series "Revenue:10,20,15,30" --title "Revenue"

    # Show presentation info
    python -m powerpoint_toolkit info -f slides.pptx
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import List, Optional


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="powerpoint_toolkit",
        description="PowerPoint Automation Toolkit — create and modify .pptx files.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # ------------------------------------------------------------------ create
    p_create = subparsers.add_parser(
        "create",
        help="Create a new, empty .pptx file.",
    )
    p_create.add_argument(
        "-o", "--output",
        required=True,
        metavar="FILE",
        help="Output .pptx file path.",
    )
    p_create.add_argument("--width", type=float, default=13.33, metavar="INCHES")
    p_create.add_argument("--height", type=float, default=7.5, metavar="INCHES")

    # --------------------------------------------------------------- add-slide
    p_slide = subparsers.add_parser(
        "add-slide",
        help="Append a slide to an existing presentation.",
    )
    p_slide.add_argument("-f", "--file", required=True, metavar="FILE")
    p_slide.add_argument("-o", "--output", metavar="FILE",
                         help="Output file (defaults to overwriting --file).")
    p_slide.add_argument(
        "--layout",
        default="blank",
        choices=[
            "blank", "title", "title_and_content", "title_only",
            "two_content", "comparison", "content_with_caption",
            "picture_with_caption", "section_header",
        ],
        metavar="LAYOUT",
    )
    p_slide.add_argument("--title", metavar="TEXT")
    p_slide.add_argument("--subtitle", metavar="TEXT")
    p_slide.add_argument("--background", metavar="HEX", help="Background hex color, e.g. #1F3864")

    # ---------------------------------------------------------------- add-text
    p_text = subparsers.add_parser("add-text", help="Add a text box to a slide.")
    p_text.add_argument("-f", "--file", required=True, metavar="FILE")
    p_text.add_argument("-o", "--output", metavar="FILE")
    p_text.add_argument("--slide", type=int, default=0, metavar="INDEX",
                        help="Zero-based slide index (default 0).")
    p_text.add_argument("--text", required=True, metavar="TEXT")
    p_text.add_argument("--left", type=float, default=1.0, metavar="INCHES")
    p_text.add_argument("--top", type=float, default=1.5, metavar="INCHES")
    p_text.add_argument("--width", type=float, default=8.0, metavar="INCHES")
    p_text.add_argument("--height", type=float, default=1.0, metavar="INCHES")
    p_text.add_argument("--font-size", type=float, metavar="PT")
    p_text.add_argument("--font-name", metavar="NAME")
    p_text.add_argument("--color", metavar="HEX")
    p_text.add_argument("--bold", action="store_true")
    p_text.add_argument("--italic", action="store_true")
    p_text.add_argument("--alignment", choices=["left", "center", "right", "justify"])

    # --------------------------------------------------------------- add-image
    p_img = subparsers.add_parser("add-image", help="Add an image to a slide.")
    p_img.add_argument("-f", "--file", required=True, metavar="FILE")
    p_img.add_argument("-o", "--output", metavar="FILE")
    p_img.add_argument("--slide", type=int, default=0, metavar="INDEX")
    p_img.add_argument("--image", required=True, metavar="PATH")
    p_img.add_argument("--left", type=float, default=1.0, metavar="INCHES")
    p_img.add_argument("--top", type=float, default=1.5, metavar="INCHES")
    p_img.add_argument("--width", type=float, metavar="INCHES")
    p_img.add_argument("--height", type=float, metavar="INCHES")

    # --------------------------------------------------------------- add-table
    p_tbl = subparsers.add_parser("add-table", help="Add a table to a slide.")
    p_tbl.add_argument("-f", "--file", required=True, metavar="FILE")
    p_tbl.add_argument("-o", "--output", metavar="FILE")
    p_tbl.add_argument("--slide", type=int, default=0, metavar="INDEX")
    p_tbl.add_argument(
        "--rows", action="append", required=True, metavar="CSV",
        help="Comma-separated row values.  Repeat for multiple rows.",
    )
    p_tbl.add_argument("--no-header", action="store_true",
                       help="Treat first row as data (not a header).")
    p_tbl.add_argument("--left", type=float, default=1.0, metavar="INCHES")
    p_tbl.add_argument("--top", type=float, default=1.5, metavar="INCHES")
    p_tbl.add_argument("--width", type=float, default=8.0, metavar="INCHES")
    p_tbl.add_argument("--height", type=float, default=3.0, metavar="INCHES")
    p_tbl.add_argument("--font-size", type=float, metavar="PT")
    p_tbl.add_argument("--header-bg", metavar="HEX")
    p_tbl.add_argument("--header-color", metavar="HEX")
    p_tbl.add_argument("--alt-row-color", metavar="HEX")

    # --------------------------------------------------------------- add-chart
    p_chart = subparsers.add_parser("add-chart", help="Add a chart to a slide.")
    p_chart.add_argument("-f", "--file", required=True, metavar="FILE")
    p_chart.add_argument("-o", "--output", metavar="FILE")
    p_chart.add_argument("--slide", type=int, default=0, metavar="INDEX")
    p_chart.add_argument(
        "--type",
        default="column",
        choices=[
            "bar", "bar_stacked", "column", "column_stacked",
            "line", "line_markers", "pie", "doughnut", "area", "scatter",
        ],
        metavar="TYPE",
    )
    p_chart.add_argument(
        "--categories", required=True, metavar="CSV",
        help="Comma-separated category labels.",
    )
    p_chart.add_argument(
        "--series", action="append", required=True, metavar="NAME:V1,V2,...",
        help="Series in the form Name:v1,v2,...  Repeat for multiple series.",
    )
    p_chart.add_argument("--title", metavar="TEXT")
    p_chart.add_argument("--no-legend", action="store_true")
    p_chart.add_argument("--data-labels", action="store_true")
    p_chart.add_argument("--left", type=float, default=1.0, metavar="INCHES")
    p_chart.add_argument("--top", type=float, default=1.5, metavar="INCHES")
    p_chart.add_argument("--width", type=float, default=8.0, metavar="INCHES")
    p_chart.add_argument("--height", type=float, default=4.5, metavar="INCHES")

    # -------------------------------------------------------------------- info
    p_info = subparsers.add_parser("info", help="Print presentation metadata.")
    p_info.add_argument("-f", "--file", required=True, metavar="FILE")
    p_info.add_argument("--json", action="store_true", dest="as_json",
                        help="Output as JSON.")

    return parser


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------

def cmd_create(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    prs = Presentation(width_inches=args.width, height_inches=args.height)
    saved = prs.save(args.output)
    print(f"Created: {saved}")


def cmd_add_slide(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    prs = Presentation(args.file)
    slide = prs.add_slide(layout=args.layout)

    if args.background:
        slide.set_background_color(args.background)
    if args.title:
        slide.set_title(args.title)
    if args.subtitle:
        slide.set_subtitle(args.subtitle)

    out = args.output or args.file
    saved = prs.save(out)
    print(f"Saved ({prs.slide_count} slide(s)): {saved}")


def cmd_add_text(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    prs = Presentation(args.file)
    slide = prs.get_slide(args.slide)
    slide.add_text_box(
        args.text,
        left=args.left,
        top=args.top,
        width=args.width,
        height=args.height,
        font_size=args.font_size,
        font_name=args.font_name,
        bold=True if args.bold else None,
        italic=True if args.italic else None,
        color=args.color,
        alignment=args.alignment,
    )
    out = args.output or args.file
    saved = prs.save(out)
    print(f"Saved: {saved}")


def cmd_add_image(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    prs = Presentation(args.file)
    slide = prs.get_slide(args.slide)
    slide.add_image(
        args.image,
        left=args.left,
        top=args.top,
        width=args.width,
        height=args.height,
    )
    out = args.output or args.file
    saved = prs.save(out)
    print(f"Saved: {saved}")


def cmd_add_table(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    data = [row.split(",") for row in args.rows]
    prs = Presentation(args.file)
    slide = prs.get_slide(args.slide)
    slide.add_table(
        data,
        has_header=not args.no_header,
        left=args.left,
        top=args.top,
        width=args.width,
        height=args.height,
        font_size=args.font_size,
        header_bg_color=args.header_bg,
        header_font_color=args.header_color,
        alternate_row_color=args.alt_row_color,
    )
    out = args.output or args.file
    saved = prs.save(out)
    print(f"Saved: {saved}")


def cmd_add_chart(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    categories = [c.strip() for c in args.categories.split(",")]
    series = []
    for s in args.series:
        name, values_str = s.split(":", 1)
        values = [float(v.strip()) for v in values_str.split(",")]
        series.append((name.strip(), values))

    prs = Presentation(args.file)
    slide = prs.get_slide(args.slide)
    slide.add_chart(
        args.type,
        categories,
        series,
        title=args.title,
        left=args.left,
        top=args.top,
        width=args.width,
        height=args.height,
        has_legend=not args.no_legend,
        has_data_labels=args.data_labels,
    )
    out = args.output or args.file
    saved = prs.save(out)
    print(f"Saved: {saved}")


def cmd_info(args: argparse.Namespace) -> None:
    from powerpoint_toolkit.presentation import Presentation

    prs = Presentation(args.file)
    info = prs.info()
    info["slides"] = []
    for i, slide in enumerate(prs.slides):
        info["slides"].append(
            {"index": i, "shapes": len(slide.shapes)}
        )
    if args.as_json:
        print(json.dumps(info, indent=2))
    else:
        print(f"File:   {info['path']}")
        print(f"Slides: {info['slide_count']}")
        print(f"Size:   {info['slide_width_inches']}\" × {info['slide_height_inches']}\"")
        for s in info["slides"]:
            print(f"  Slide {s['index']}: {s['shapes']} shape(s)")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

_COMMANDS = {
    "create": cmd_create,
    "add-slide": cmd_add_slide,
    "add-text": cmd_add_text,
    "add-image": cmd_add_image,
    "add-table": cmd_add_table,
    "add-chart": cmd_add_chart,
    "info": cmd_info,
}


def main(argv: Optional[List[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    handler = _COMMANDS.get(args.command)
    if handler is None:
        parser.print_help()
        return 1
    try:
        handler(args)
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
