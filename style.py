from abc import ABC

from helpers import FontStyle


class BaseStyle(ABC):
    title: FontStyle
    field: FontStyle
    body: FontStyle
    marginals: FontStyle


class DefaultStyle(BaseStyle):
    title = FontStyle("Calibri", "bold", 22)
    field = FontStyle("Calibri", "bold", 12)
    body = FontStyle("Calibri", "normal", 11)
    marginals = FontStyle("Calibri", "normal", 12)


default_chan_style = FontStyle("Calibri", "bold", 18)
default_position_style = FontStyle("Calibri", "bold", 18)
