"""Shape and text-formatting utilities."""

from __future__ import annotations

from typing import Optional, Tuple

from pptx.dml.color import RGBColor
from pptx.util import Pt


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert a CSS hex colour string to a :class:`~pptx.dml.color.RGBColor`.

    Accepts both ``#RRGGBB`` and ``RRGGBB`` forms.

    Examples
    --------
    >>> hex_to_rgb("#FF5733")
    RGBColor(0xff, 0x57, 0x33)
    """
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        raise ValueError(f"Invalid hex colour {hex_color!r}.  Expected 6 hex digits.")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


def apply_text_format(
    run,
    *,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    font_size: Optional[float] = None,
    font_name: Optional[str] = None,
    color: Optional[str] = None,
) -> None:
    """Apply formatting attributes to a python-pptx *run* object in place.

    Parameters
    ----------
    run:
        A ``pptx.text.text._Run`` instance.
    bold, italic, underline:
        Toggle the respective style.
    font_size:
        Font size in *points*.
    font_name:
        Font family name, e.g. ``"Calibri"``.
    color:
        Hex colour string, e.g. ``"#FF0000"``.
    """
    font = run.font
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if underline is not None:
        font.underline = underline
    if font_size is not None:
        font.size = Pt(font_size)
    if font_name is not None:
        font.name = font_name
    if color is not None:
        font.color.rgb = hex_to_rgb(color)


def apply_paragraph_format(
    paragraph,
    *,
    alignment: Optional[str] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    line_spacing: Optional[float] = None,
) -> None:
    """Apply paragraph-level formatting to a python-pptx *paragraph* object.

    Parameters
    ----------
    paragraph:
        A ``pptx.text.text._Paragraph`` instance.
    alignment:
        One of ``"left"``, ``"center"``, ``"right"``, ``"justify"``.
    space_before:
        Space before the paragraph in points.
    space_after:
        Space after the paragraph in points.
    line_spacing:
        Line spacing in points.
    """
    from pptx.enum.text import PP_ALIGN  # local import to avoid circular

    _align_map = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }
    if alignment is not None:
        key = alignment.lower()
        if key not in _align_map:
            raise ValueError(
                f"Unknown alignment {alignment!r}.  "
                f"Choose from: {', '.join(_align_map)}."
            )
        paragraph.alignment = _align_map[key]

    if space_before is not None or space_after is not None or line_spacing is not None:
        pf = paragraph.paragraph_format
        if space_before is not None:
            pf.space_before = Pt(space_before)
        if space_after is not None:
            pf.space_after = Pt(space_after)
        if line_spacing is not None:
            pf.line_spacing = Pt(line_spacing)


def parse_position(
    left: float,
    top: float,
    width: float,
    height: float,
) -> Tuple[int, int, int, int]:
    """Convert inch measurements to EMUs (the unit used by python-pptx).

    Returns a tuple of ``(left, top, width, height)`` in EMU.
    """
    from pptx.util import Inches

    return (
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
