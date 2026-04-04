"""Useful helpers and dataclasses for paperwork generation."""

import datetime
import logging
import re
from dataclasses import dataclass
from decimal import Decimal
from typing import List, Optional, Self, Union

import openpyxl

logger = logging.getLogger(__name__)


@dataclass
class ShowData:
    """Dataclass for storing information about the show."""

    show_name: Optional[str] = None
    ld_name: Optional[str] = None
    revision: Optional[str] = None
    date: datetime.datetime = datetime.datetime.now()

    def print_date(self) -> str:
        """Returns the date in YYYY/MM/DD form"""
        return self.date.strftime("%Y/%m/%d")

    def generate_slug(self, title: str = "Paperwork") -> str:
        """Generate a filename slug from the show information"""
        if self.show_name is None or self.revision is None:
            logger.info("Not enough show data to make a nice output filename, using default")
            return title
        return f"{self.show_name.replace(' ', '')}_{title}_" + re.sub(r"\W+", "", self.revision)


@dataclass
class InstrumentPower:
    """Dataclass for instrument power and formatting"""

    """
    Looks for power strings within strings.
    ex. `INSTR 300W` or `INSTR 300 kW` but not `INSTR 300`
    """
    POWER_REGEX = r"\d+\.?\d*\s*[kMGm]?W"

    def __init__(self, input_string: Union[str | int | float | None]):
        if input_string is None:
            self.power: Decimal = Decimal(0)
        elif isinstance(input_string, (float, int)):
            self.power: Decimal = Decimal(input_string)
        elif isinstance(input_string, str):
            if input_string.strip() == "":
                self.power: Decimal = Decimal(0)
            else:
                # Remove non-numeric/decimal values
                self.power: Decimal = Decimal(re.sub(r"[^\d\.\-]", "", input_string))
                if "k" in input_string:
                    self.power *= 1000
                if "M" in input_string:
                    self.power *= 1000 * 1000

        self.power = self.power.normalize()
        if self.power < 0:
            raise ValueError(f"Light sucker detected ({self.power} < 0W)")

    def format(self):
        # Remove sigfigs but prints more "cleanly"
        powerval = (
            self.power.quantize(Decimal(1))
            if self.power == self.power.to_integral()
            else self.power.normalize()
        )
        if self.power >= 1000:
            return f"{powerval / 1000}kW"
        return f"{powerval}W"


@dataclass
class DMXAddress:
    """Dataclass for DMX address and formatting"""

    def __init__(self, input_string: Union[str | int]):
        if isinstance(input_string, int):
            # Assume absolute address
            if input_string <= 0:
                raise ValueError(f"Absolute address cannot be less than 1 (got {input_string})")
            self.absolute_address: int = input_string
        elif input_string.isdigit():
            if int(input_string) <= 0:
                raise ValueError(f"Absolute address cannot be less than 1 (got {input_string})")
            self.absolute_address: int = int(input_string)
        elif "/" in input_string or ":" in input_string:
            # universe : or / relative address
            if "/" in input_string:
                universe, rel_addr = input_string.split("/", 2)
            else:
                universe, rel_addr = input_string.split(":", 2)

            # Will raise ValueError if failure
            universe_int = int(universe)
            rel_addr_int = int(rel_addr)

            if universe_int <= 0:
                raise ValueError(f"Universe cannot be less than 1 (got {universe_int})")
            if rel_addr_int <= 0:
                raise ValueError(f"Relative address cannot be less than 1 (got {rel_addr_int})")
            if rel_addr_int > 512:
                raise ValueError(
                    f"Relative address cannot be greater than 512 (got {rel_addr_int})"
                )

            self.absolute_address: int = (universe_int - 1) * 512 + rel_addr_int
        else:
            raise ValueError(f"Unclear address {input_string}")

    def get_universe(self) -> int:
        return int(((self.absolute_address - 1) / 512) + 1)

    def get_relative_address(self) -> int:
        return int(((self.absolute_address - 1) % 512) + 1)

    def format_slash(self) -> str:
        return f"{self.get_universe()}/{self.get_relative_address()}"

    def format_colon(self) -> str:
        return f"{self.get_universe()}:{self.get_relative_address()}"

    def format_slash_conditional(self) -> str:
        """If first universe, don't add slash"""
        if self.absolute_address < 513:
            return f"{self.absolute_address}"
        return self.format_slash()


@dataclass
class Gel:
    """Dataclass for storing information about a gel."""

    name: str
    name_sort: str
    company: str

    # TODO: add nice name from lighting_filters

    @classmethod
    def parse_name(cls, gel: str) -> Self:
        """Returns a Gel from a common name (ex. R355 or L201)."""
        gel = gel.strip()
        if re.search(r"^AP\d+$", gel, re.IGNORECASE):
            company = "Apollo"
        elif re.search(r"^G\d+$", gel, re.IGNORECASE):
            company = "GAM"
        elif re.search(r"^L\d+$", gel, re.IGNORECASE):
            company = "Lee"
        elif re.search(r"^R\d+$", gel, re.IGNORECASE):
            company = "Rosco"
        else:
            logger.warning("Unknown company prefix for gel %s", gel)
            return cls(gel, gel, "")

        # consistent formatting for known gels
        gel = gel.upper()

        if company == "Rosco" and re.match(r"^R3\d\d$", gel, re.IGNORECASE):
            # Rosco extended gel, this is basically a .5 gel for sorting purposes
            gelsort = "R" + gel[2:] + ".3"
        else:
            gelsort = gel

        return cls(gel, gelsort, company)

    @classmethod
    def parse_gel_string(cls, gel: str) -> List[Self]:
        """Parses a complex gel string (ex. L202x2 + R119 + R26)"""
        gel = gel.strip()
        if gel is None or gel == "":
            return [cls("", "", "")]

        gel_list = []
        # Only supports + as a separator for now
        for i in gel.split("+"):
            i = i.strip()
            if len(i.split("x")) > 1:
                # Repeat gel situation (ex. L201x2)
                g, count = i.split("x")
                for _ in range(0, int(count)):
                    gel_list.append(cls.parse_name(g))
            else:
                gel_list.append(cls.parse_name(i))

        return gel_list


def parse_frame_size(frame_str: str) -> str:
    """Parses a frame size."""
    # TODO: make this generic with a unit conversion program
    frame_str = frame_str.strip()
    if len(frame_str.lower().split("x")) > 1:
        # rectangular frame!
        lengths = []
        for i in frame_str.lower().split("x"):
            lengths.append(parse_frame_size(i))

        ret_string = ""
        for idx, val in enumerate(lengths):
            if idx + 1 < len(lengths):
                ret_string += f"{val}x"
            else:
                ret_string += f"{val}"

        return ret_string

    if frame_str == "":
        return ""

    if "cm" in frame_str:
        frame_size = frame_str.replace("cm", "").strip()
        return f"{frame_size} cm"

    if "'" in frame_str:
        # what in the sky cyc
        if '"' in frame_str:
            raise ValueError(f"Mixed dimensions unsupported ({frame_str})")

        frame_size = float(frame_str.replace("'", "")) * 12
        return f'{frame_size}"'

    frame_size = frame_str.replace('"', "").replace("in", "").strip()
    return f'{frame_size}"'


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
            return openpyxl.styles.Font(name=self.font_family, size=self.font_size, bold=True)

        if self.font_weight == "normal":
            return openpyxl.styles.Font(name=self.font_family, size=self.font_size, bold=False)

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
