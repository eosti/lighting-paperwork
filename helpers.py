"""Useful helpers and dataclasses for paperwork generation."""

import datetime
import re
from dataclasses import dataclass


@dataclass
class ShowData:
    """Dataclass for storing information about the show."""

    show_name: str
    ld_name: str
    revision: str
    date: datetime = datetime.datetime.now()

    def print_date(self) -> str:
        """Returns the date in YYYY/MM/DD form"""
        return self.date.strftime("%Y/%m/%d")


@dataclass
class Gel:
    """Dataclass for storing information about a gel."""

    name: str
    name_sort: str
    company: str

    @classmethod
    def parse_name(cls, gel: str):
        """Returns a Gel from a common name (ex. R355 or L201)."""
        if gel.startswith("AP"):
            company = "Apollo"
        elif gel.startswith("G"):
            company = "GAM"
        elif gel.startswith("L"):
            company = "Lee"
        elif gel.startswith("R"):
            company = "Rosco"
        else:
            raise ValueError(f"Unknown company prefix for gel {gel}")

        gelsort = gel
        if company == "Rosco":
            if re.match(r"^R3\d\d$", gel):
                # Rosco extended gel, this is basically a .5 gel
                gelsort = "R" + gel[2:] + ".3"

        return cls(gel, gelsort, company)


@dataclass
class FontStyle:
    """Dataclass for storing CSS font style information."""

    font_family: str
    font_weight: str
    font_size: int

    def to_css(self) -> str:
        """Returns a CSS string with the font information."""
        return (
            f"font-family: {self.font_family}; "
            f"font-weight: {self.font_weight}; font-size: {self.font_size}pt; "
        )

    def span(self, body: str, style: str = "") -> str:
        """Returns a `span` element formatted with the font information."""
        return f"<span style='{self.to_css()}{style}'>{body}</span>"

    def p(self, body: str, style: str = "") -> str:
        """Returns a `p` element formatted with the font information."""
        return f"<p style='{self.to_css()}{style}'>{body}</p>"
