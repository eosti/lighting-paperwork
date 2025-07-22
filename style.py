"""Styling information for paperwork"""

from abc import ABC
from dataclasses import dataclass

from helpers import FontStyle


@dataclass
class BaseStyle(ABC):
    """Class to store font style for paperwork"""

    title: FontStyle
    field: FontStyle
    body: FontStyle
    marginals: FontStyle


@dataclass
class DefaultStyle(BaseStyle):
    """Default style for all paperwork"""

    title = FontStyle("Calibri", "bold", 22)
    field = FontStyle("Calibri", "bold", 12)
    body = FontStyle("Calibri", "normal", 11)
    marginals = FontStyle("Calibri", "normal", 12)


default_chan_style = FontStyle("Calibri", "bold", 18)
default_position_style = FontStyle("Calibri", "bold", 18)
