"""PowerPoint Automation Toolkit.

A Python toolkit for creating and manipulating PowerPoint presentations
programmatically using python-pptx.

Quick start::

    from powerpoint_toolkit import Presentation

    prs = Presentation()
    slide = prs.add_slide(layout="title")
    slide.set_title("Hello, World!")
    slide.set_subtitle("Built with powerpoint-automation-toolkit")
    prs.save("hello.pptx")
"""

from powerpoint_toolkit.presentation import Presentation
from powerpoint_toolkit.slides import Slide
from powerpoint_toolkit.charts import ChartBuilder

__all__ = ["Presentation", "Slide", "ChartBuilder"]
__version__ = "1.0.0"
