"""Useful helpers and dataclasses for paperwork generation."""

import datetime
import logging
import re
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Self

import openpyxl.styles as openpyxl_styles

logger = logging.getLogger(__name__)


@dataclass
class ShowData:
    """Dataclass for storing information about the show."""

    show_name: str | None = None
    ld_name: str | None = None
    revision: str | None = None
    rev_date: datetime.datetime = field(default_factory=lambda: datetime.datetime.now(datetime.UTC))

    def print_date(self) -> str:
        """Return the stored date in YYYY/MM/DD form."""
        return self.rev_date.astimezone().strftime("%Y/%m/%d")

    def generate_slug(self, title: str = "Paperwork") -> str:
        """Generate a filename slug from the show information."""
        if self.show_name is None or self.revision is None:
            logger.info("Not enough show data to make a nice output filename, using default")
            return title
        return f"{self.show_name.replace(' ', '')}_{title}_" + re.sub(r"\W+", "", self.revision)


@dataclass
class InstrumentPower:
    """Dataclass for instrument power and formatting.

    Attributes:
        POWER_REGEX: A regex that looks for power strings within strings.
            ex. matches `INSTR 300W` or `INSTR 300 kW` but not `INSTR 300`.
        power: Numeric representation of the instrument's power.

    """

    POWER_REGEX = r"\d+\.?\d*\s*[kMGm]?W"

    def __init__(self, input_string: str | float | None) -> None:
        """Initialize class with unformatted data.

        Args:
            input_string: A string or number that somewhat resembles a valid power.

        """
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

    def format(self) -> str:
        """Print power in a consistent format as XX[.X]W."""
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
    """Dataclass for DMX address and formatting.

    Attributes:
        absolute_address: The absolute DMX address, 1-indexed.

    """

    def __init__(self, input_string: str | int) -> None:
        """Initialize class with address.

        Args:
            input_string: The DMX address as a string or number.
                May be an absolute address or slash/colon-formatted.

        Raises:
            ValueError: The input_string does not resemble a valid DMX address.

        """
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
        """Determine the universe that the address belongs in."""
        return int(((self.absolute_address - 1) / 512) + 1)

    def get_relative_address(self) -> int:
        """Determine the relative address within the universe."""
        return int(((self.absolute_address - 1) % 512) + 1)

    def format_slash(self) -> str:
        """Return a string in the form of universe/rel_address."""
        return f"{self.get_universe()}/{self.get_relative_address()}"

    def format_colon(self) -> str:
        """Return a string in the form of universe:rel_address."""
        return f"{self.get_universe()}:{self.get_relative_address()}"

    def format_slash_conditional(self) -> str:
        """Like format_slash, but if in the first universe don't add the `1/`."""
        if self.absolute_address < 513:
            return f"{self.absolute_address}"
        return self.format_slash()


@dataclass
class Gel:
    """Dataclass for storing information about a gel.

    Attributes:
        name: The name code for a gel (ex. R34 or AP4400)
        name_sort: typically the same as `name` but may be slightly altered for better sorting
            ex. for Rosco 300-series gels, name_sort will be a .3 value (R334 -> R34.3)
        company: The name of the manufacturer of the gel.

    """

    name: str
    name_sort: str
    company: str

    # TODO(eosti): add nice name from lighting_filters
    # https://github.com/eosti/lighting-paperwork/issues/9

    @classmethod
    def _parse_name(cls, gel: str) -> Self:
        """Return a Gel from a common name.

        Args:
            gel: name code for a gel (ex. R355 or L201).

        """
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
    def parse_gel(cls, gel: str) -> list[Self]:
        """Parse a complex gel string into a list of gels.

        Args:
            gel: a complex gel string (ex. 'L202x2 + R119 + R26')

        Returns:
            A list of Gels, one entry per physical gel.
                ex. 'L202x2 + R119 + R26' would return [Gel(L202), Gel(L202), Gel(R119), Gel(R26)].

        """
        gel = gel.strip()
        if gel is None or gel == "":
            return [cls("", "", "")]

        gel_list = []
        # Only supports + as a separator for now
        for i in gel.split("+"):
            gel_name = i.strip()
            if len(gel_name.split("x")) > 1:
                # Repeat gel situation (ex. L201x2)
                g = gel_name.split("x")
                gel_list.extend(cls._parse_name(g[0]) for x in range(int(g[1])))
            else:
                gel_list.append(cls._parse_name(gel_name))

        return gel_list


def parse_frame_size(frame_str: str) -> str:
    """Parse a frame size.

    Args:
        frame_str: a string with a single or two dimensions. Assumes inches unless otherwise stated.
            Valid units: " or in for inches, ' for feet, cm for centimeters.
            Ex. `6.25"`, `12"x14"`, `7.5`

    Returns:
        Consistently formatted dimension string.
        Imperial units will be converted to inches and formatted as XX"

    """
    # TODO(eosti): make this generic with a unit conversion program
    # https://github.com/eosti/lighting-paperwork/issues/10

    frame_str = frame_str.strip()
    if len(frame_str.lower().split("x")) > 1:
        # rectangular frame!
        lengths = []
        lengths.extend(parse_frame_size(x) for x in frame_str.lower().split("x"))

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
    """Dataclass for storing CSS font style information.

    Attributes:
        font_family: The PostScript name of a font family installed locally
        font_weight: The CSS weight of the font (100-900). May also use relative values.
        font_size: The size of the font, in pt.

    """

    font_family: str
    font_weight: str
    font_size: int

    def to_css(self) -> str:
        """Return a CSS string with the font information."""
        return (
            f"font-family: {self.font_family}; "
            f"font-weight: {self.font_weight}; font-size: {self.font_size}pt; "
        )

    def span(self, body: str, style: str = "") -> str:
        """Return a `span` element formatted with the font information."""
        return f"<span style='{self.to_css()}{style}'>{body}</span>"

    def p(self, body: str, style: str = "") -> str:
        """Return a `p` element formatted with the font information."""
        return f"<p style='{self.to_css()}{style}'>{body}</p>"

    def excel(self) -> openpyxl_styles.Font:
        """Return an openpyxl Style with the selected font.

        Note that only `normal` and `bold` font weights are permitted.
        """
        if self.font_weight == "bold":
            return openpyxl_styles.Font(name=self.font_family, size=self.font_size, bold=True)

        if self.font_weight == "normal":
            return openpyxl_styles.Font(name=self.font_family, size=self.font_size, bold=False)

        raise ValueError(f"Unsupported weight {self.font_weight}")


@dataclass
class StyledContent:
    """Dataclass for storing content/style pairs.

    Attributes:
        content: An HTML string
        style: A CSS string to format `content`

    """

    content: str = ""
    style: str = ""


@dataclass
class FormattingQuirks:
    """Collection of formatting differences between the various export formats.

    Attributes:
        empty_str: What string to represent an empty cell
        hidden_fmt: What CSS "color" should something be in order to be hidden?

    """

    empty_str: str
    hidden_fmt: str


html_quirks = FormattingQuirks("&nbsp;", "transparent")
excel_quirks = FormattingQuirks("", "#FFFFFF")
