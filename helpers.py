import datetime
import re
from dataclasses import dataclass
from typing import Optional


@dataclass
class ShowData:
    show_name: str
    ld_name: str
    revision: str
    date: datetime = datetime.datetime.now()

    def print_date(self) -> str:
        """Returns a string in YYYY/MM/DD form"""
        return self.date.strftime("%Y/%m/%d")


@dataclass
class Gel:
    name: str
    name_sort: str
    company: str

    @classmethod
    def parse_name(cls, gel: str):
        if gel.startswith("AP"):
            company = "Apollo"
        elif gel.startswith("G"):
            company = "GAM"
        elif gel.startswith("L"):
            company = "Lee"
        elif gel.startswith("R"):
            company = "Rosco"

        gelsort = gel
        if company == "Rosco":
            if re.match(r"^R3\d\d$", gel):
                # Rosco extended gel, this is basically a .5 gel
                gelsort = "R" + gel[2:] + ".3"

        return cls(gel, gelsort, company)


@dataclass
class FontStyle:
    font_family: str
    font_weight: str
    font_size: int

    def to_css(self) -> str:
        return f"font-family: {self.font_family}; font-weight: {self.font_weight}; font-size: {self.font_size}pt; "

    def span(self, body: str, style: str = "") -> str:
        return f"<span style='{self.to_css()}{style}'>{body}</span>"

    def p(self, body: str, style: str = "") -> str:
        return f"<p style='{self.to_css()}{style}'>{body}</p>"


