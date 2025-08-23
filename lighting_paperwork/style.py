"""Styling information for paperwork"""

from dataclasses import dataclass

from .helpers import FontStyle


@dataclass
class BaseStyle:
    """Class to store font style for paperwork"""

    title: FontStyle
    field: FontStyle
    body: FontStyle
    marginals: FontStyle


default_style = BaseStyle(
    title=FontStyle("Calibri", "bold", 22),
    field=FontStyle("Calibri", "bold", 12),
    body=FontStyle("Calibri", "normal", 11),
    marginals=FontStyle("Calibri", "normal", 12),
)

default_chan_style = FontStyle("Calibri", "bold", 18)
default_position_style = FontStyle("Calibri", "bold", 18)
