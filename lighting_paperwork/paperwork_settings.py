"""Module for defining all settings relating to paperwork generation."""

import datetime
import logging
import re
from dataclasses import dataclass
from typing import Annotated, Literal

import openpyxl.styles as openpyxl_styles
from pydantic import AliasChoices, BaseModel, Field, FilePath, StringConstraints
from pydantic_settings import (
    BaseSettings,
    CliPositionalArg,
    SettingsConfigDict,
)

logger = logging.getLogger(__name__)


class ShowData(BaseModel):
    """Model for storing information about the show."""

    show_name: str | None = Field(default=None, description="Show name")
    ld_name: str | None = Field(default=None, description="Lighting designer initials")
    revision: str | None = Field(default=None, description="Revision string (ex. 'Rev. A')")
    rev_date: datetime.datetime = Field(default_factory=lambda: datetime.datetime.now(datetime.UTC))

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


class PaperworkSettings(BaseSettings):
    """Top level settings schema."""

    model_config = SettingsConfigDict(
        cli_parse_args=True,
        cli_avoid_json=True,
        cli_hide_none_type=True,
        cli_shortcuts={
            "show_info.show_name": "show",
            "show_info.ld_name": "ld",
            "show_info.revision": "rev",
        },
    )
    input_file: CliPositionalArg[FilePath] = Field(
        validation_alias=AliasChoices("file"), description="CSV or XML from Vectorworks"
    )
    show_info: ShowData = ShowData()
    log_level: Annotated[
        Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"], StringConstraints(to_upper=True)
    ] = Field(
        default="INFO",
        validation_alias=AliasChoices("loglevel"),
        description="Change the log level",
    )
    output_type: Literal["pdf", "html", "csv"] = Field(
        default="pdf",
        validation_alias=AliasChoices("out"),
        description="Choose the output file type",
    )
