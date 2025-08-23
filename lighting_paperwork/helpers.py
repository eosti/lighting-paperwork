"""Useful helpers and dataclasses for paperwork generation."""

import datetime
import re
from dataclasses import dataclass
from typing import Self

import openpyxl


@dataclass
class ShowData:
    """Dataclass for storing information about the show."""

    show_name: str
    ld_name: str
    revision: str
    date: datetime.datetime = datetime.datetime.now()

    def print_date(self) -> str:
        """Returns the date in YYYY/MM/DD form"""
        return self.date.strftime("%Y/%m/%d")

    def generate_slug(self, title: str = "Paperwork") -> str:
        """Generate a filename slug from the show information"""
        return f"{self.show_name.replace(' ', '')}_{title}_" + re.sub(
            r"\W+", "", self.revision
        )


@dataclass
class Gel:
    """Dataclass for storing information about a gel."""

    name: str
    name_sort: str
    company: str

    @classmethod
    def parse_name(cls, gel: str) -> Self:
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

    def excel(self) -> openpyxl.styles.Font:
        if self.font_weight == "bold":
            return openpyxl.styles.Font(
                name=self.font_family, size=self.font_size, bold=True
            )

        if self.font_weight == "normal":
            return openpyxl.styles.Font(
                name=self.font_family, size=self.font_size, bold=False
            )

        raise ValueError(f"Unsupported weight {self.font_weight}")


@dataclass
class FormattingQuirks:
    """
    Collection of formatting differences between the
    various export formats.
    """

    """What string to represent an empty cell"""
    empty_str: str
    """What argument to CSS `color:` to hide text"""
    hidden_fmt: str


html_quirks = FormattingQuirks("&nbsp;", "transparent")
excel_quirks = FormattingQuirks("", "#FFFFFF")
