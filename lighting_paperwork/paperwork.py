"""Base paperwork generation class."""

import logging
import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import NotRequired, Self, TypedDict, Unpack

import numpy as np
import openpyxl
import pandas as pd
from pandas.io.formats.style import Styler
from pandas.io.formats.style_render import CSSDict

from lighting_paperwork import excel_formatter
from lighting_paperwork.helpers import (
    DMXAddress,
    FontStyle,
    FormattingQuirks,
    InstrumentPower,
    ShowData,
    StyledContent,
    excel_quirks,
    html_quirks,
)
from lighting_paperwork.style import BaseStyle, default_style

logger = logging.getLogger(__name__)


class StyleDataParams(TypedDict):
    """**kwargs for style_data functions."""

    body_style: FontStyle
    col_width: list[int]
    border_weight: float
    quirks: FormattingQuirks
    chan_style: NotRequired[FontStyle]


class StyleFieldParams(TypedDict):
    """**kwargs for style_field functions."""

    header_style: FontStyle
    col_width: list[int]
    border_weight: float


class PaperworkGenerator(ABC):
    """Base paperwork generation class.

    Attributes:
        (should be wrapped up in a Config dataclass eventually)

    """

    def __init__(
        self,
        vw_export: pd.DataFrame,
        show_data: ShowData | None = None,
        style: BaseStyle = default_style,
        border_weight: float = 1.0,
    ) -> None:
        """Set class vars for data and style."""
        self.vw_export = vw_export
        self.df = self.vw_export.copy()
        self.show_data = show_data
        self.style = style
        # 1px doesn't render right on Firefox, use 1.5px min to workaround.
        self.border_weight = border_weight

    def set_show_data(self, show_name: str, ld_name: str, revision: str) -> None:
        """Save show data for later use."""
        self.show_data = ShowData(show_name=show_name, ld_name=ld_name, revision=revision)

    display_name: str
    col_widths: tuple[int, ...]
    page_width: int = 100
    formatting_quirks = html_quirks
    no_color_text = "N/C"

    @abstractmethod
    def generate_df(self) -> Self:
        """Generate a DataFrame with sorted information.

        Using `self.vw_export`, generate a DataFrame that contains the
        necessary sorted information for the paperwork type.
        """

    @staticmethod
    @abstractmethod
    def style_data(df: pd.DataFrame, /, **kwargs: Unpack[StyleDataParams]) -> pd.DataFrame:
        """Styles the data for a table."""

    @staticmethod
    @abstractmethod
    def style_fields(index: pd.Series, /, **kwargs: Unpack[StyleFieldParams]) -> list[str]:
        """Style the fields (i.e. headers) for a table."""

    def _make_common(self) -> pd.io.formats.style.Styler:
        """Run common make tasks for html and excel."""
        self.generate_df()

        styled = Styler.from_custom_template(
            str(Path(__file__).parent / "templates"), "header_footer.tpl"
        )(self.df)
        styled = styled.apply(
            type(self).style_data,
            axis=None,
            body_style=self.style.body,
            col_width=self.col_widths,
            border_weight=self.border_weight,
            quirks=self.formatting_quirks,
        )
        styled = styled.hide()
        styled = styled.apply_index(
            type(self).style_fields,  # type: ignore[reportArgumentType]
            header_style=self.style.field,
            col_width=self.col_widths,
            border_weight=self.border_weight,
            axis=1,
        )

        return styled  # noqa: RET504

    def make_html(self) -> str:
        """Generate a formatted HTML table from the generated DataFrame."""
        styled = self._make_common()

        styled = styled.set_table_attributes('class="paperwork-table"')
        styled = styled.set_table_styles(
            self.default_table_style(width=self.page_width), overwrite=False
        )

        header_html, footer_html = self.generate_header_footer(styled.uuid)  # type: ignore[reportAttributeAccessIssue]
        page_style = self.generate_page_style(
            styled.uuid,  # type: ignore[reportAttributeAccessIssue]
            "bottom-right",
            self.style.marginals.to_css(),
        )

        logger.info("Generated %s.", self.display_name)

        html = styled.to_html(
            generated_header=header_html,
            generated_footer=footer_html,
            generated_page_style=page_style,
        )
        html = self.wrap_table(html)

        return html  # noqa: RET504

    def make_excel(self, excel_path: str) -> None:
        """Add a sheet to an Excel file with the formatted DataFrame."""
        self.formatting_quirks = excel_quirks
        styled = self._make_common()

        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a") as writer:
            styled.to_excel(writer, sheet_name=self.display_name)

        wb = openpyxl.load_workbook(excel_path)
        ws = wb[self.display_name]

        # Remove index column
        ws.delete_cols(idx=1)

        # Standard formatting
        excel_formatter.add_title(ws, self.display_name, self.show_data)
        excel_formatter.page_setup(ws, 1)
        excel_formatter.set_col_widths(ws, self.col_widths, self.page_width)
        excel_formatter.wrap_all_cells(ws)
        wb.save(excel_path)

    @staticmethod
    def verify_width(width: list[int]) -> bool:
        """Verify that the col widths remain less than 100%."""
        if sum(width) > 100:
            logger.warning("Col widths too long: used %i%%!", sum(width))
            return False

        return True

    @staticmethod
    def determine_power(row: pd.Series) -> InstrumentPower:
        """Determine an instrument's power from the Instrument Type and Wattage fields.

        When the fields disagree, the values from the Wattage field takes priority.
        When either field shows 0W, the other field will take priority.

        Args:
            row: Row of a dataframe that must have a "Wattage" and a "Instrument Type" column.

        Returns:
            The determined power of an instrument.

        """
        # Collect potential powers
        wattage_field_power = InstrumentPower(row["Wattage"])
        powerstr = re.search(InstrumentPower.POWER_REGEX, row["Instrument Type"])
        if powerstr is None:
            instrument_type_power = InstrumentPower(0)
        else:
            instrument_type_power = InstrumentPower(powerstr.group())

        # Verify which we should use
        if wattage_field_power.power == 0 and instrument_type_power.power == 0:
            # No power provided
            if row["Accessory Flag"] != "1":
                logger.warning(
                    "Channel %s is infinitely efficient (%s is %s)",
                    row["Channel"],
                    row["Instrument Type"],
                    wattage_field_power.format(),
                )
            power = wattage_field_power
        elif wattage_field_power.power > 0 and instrument_type_power.power == 0:
            # There was a power in the wattage field but none parsed from instrument type
            power = wattage_field_power
        elif instrument_type_power.power > 0 and wattage_field_power.power == 0:
            # There was a power in the instrument type field but none in the wattage field
            power = instrument_type_power
        elif (
            wattage_field_power.power > 0
            and instrument_type_power.power > 0
            and wattage_field_power.power != instrument_type_power.power
        ):
            # Fields disagree, prefer the wattage field
            logger.warning(
                "Channel %s has conflicting power values (%s, %s). Using %s.",
                row["Channel"],
                wattage_field_power.format(),
                instrument_type_power.format(),
                wattage_field_power.format(),
            )
            power = wattage_field_power
        elif wattage_field_power.power == instrument_type_power.power:
            # Same power found in instrument type and wattage fields
            power = wattage_field_power
        else:
            raise ValueError("Shouldn't get here!")

        return power

    def combine_instrtype(self) -> Self:
        """Combine the Instrument Type and Power and Accessory fields into one."""
        instload = []
        for _, row in self.df.iterrows():
            power = self.determine_power(row)

            # Make sure power shows up once, after the instrument type
            if power.power == 0:
                tmp = row["Instrument Type"].strip()
            else:
                # Remove from instrument type (if existing)
                instrtype = re.sub(InstrumentPower.POWER_REGEX, "", row["Instrument Type"]).strip()
                tmp = instrtype + " " + power.format()

            # If accessory, add that here
            if row["Accessory String"] != "":
                tmp += ", " + row["Accessory String"]

            instload.append(tmp)

        # Clean up by replacing old cols with new one
        new_df = self.df.drop(["Instrument Type", "Wattage", "Accessory String"], axis=1)
        new_df["Instr Type & Load & Acc"] = instload

        self.df = new_df
        return self

    def combine_gelgobo(self) -> Self:
        """Combine the Gel and Gobo fields into one."""
        gelgobo = []
        for _, row in self.df.iterrows():
            # If no gel replace with N/C
            if row["Color"] == "":
                tmp = "" if row["Accessory Flag"] == "1" else self.no_color_text
            else:
                tmp = row["Color"]

            # Append gobo if exists
            if row["Gobo 1"] != "" and row["Gobo 2"] != "":
                tmp += ", T: " + row["Gobo 1"] + ", " + row["Gobo 2"]
            elif row["Gobo 1"] != "":
                tmp += ", T: " + row["Gobo 1"]
            elif row["Gobo 2"] != "":
                tmp += ", T: " + row["Gobo 2"]

            gelgobo.append(tmp)

        # Clean up by replacing old cols with new one
        new_df = self.df.drop(["Color", "Gobo 1", "Gobo 2"], axis=1)
        new_df["Color & Gobo"] = gelgobo
        self.df = new_df

        return self

    def format_address_slash(self) -> Self:
        """Format an absolute address into a Universe/Address string."""
        for row in self.df.itertuples():
            absaddr = int(self.df.loc[row.Index, "Absolute Address"])  # type: ignore[reportCallIssue, reportArgumentType]
            if absaddr == 0:
                # If no address set, replace it with a blank
                self.df.loc[row.Index, "Absolute Address"] = self.formatting_quirks.empty_str
            else:
                self.df.loc[row.Index, "Absolute Address"] = DMXAddress(
                    absaddr
                ).format_slash_conditional()

        slashed_df = self.df.rename(columns={"Absolute Address": "Addr"})
        self.df = slashed_df
        return self

    def repeated_index_val(self, index_str: str, df_override: pd.DataFrame | None = None) -> Self:
        """Format repeated channel numbers to use `"` to represent repeated data.

        Arguments:
            index_str: The string that represents the "main" column of the paperwork type
                ex. for a Channel Hookup, index_str = 'Chan'
            df_override: Override the use of self.df with a different dataframe

        """
        df = df_override if df_override is not None else self.df
        repeated_idx = "-1"

        prev_row: pd.Series | None = None
        for df_index, data in df.iterrows():
            if prev_row is None:
                # Initial case
                prev_row = data
                continue

            tmp_prev = None
            if data[index_str] == repeated_idx or data[index_str] == prev_row[index_str]:
                if data[index_str] == prev_row[index_str]:
                    # first repeat case
                    repeated_idx = data[index_str]

                tmp_prev = data.copy()
                for idx, val in data.items():
                    if idx == index_str:
                        # Don't "-ify the index string, just leave it blank
                        df.loc[df_index, str(idx)] = self.formatting_quirks.empty_str
                        continue
                    if index_str == "Chan" and idx == "U#":
                        # Do repeat U# to avoid confusion
                        continue
                    if val.strip() == "" or val == self.formatting_quirks.empty_str:
                        # Don't "-ify already empty fields
                        continue
                    if val == prev_row[str(idx)]:
                        df.loc[df_index, str(idx)] = '"'
            else:
                tmp_prev = data
                repeated_idx = "-1"

            prev_row = tmp_prev

        return self

    def abbreviate_col_names(self) -> Self:
        """Abbreviate common column names."""
        self.df = self.df.rename(
            columns={"Channel": "Chan", "Unit Number": "U#", "Address": "Addr"}
        )
        return self

    def verify_filter_fields(self, filter_fields: list[str]) -> None:
        """Verify certain fields exist in the dataframe."""
        for field in filter_fields:
            if field not in self.vw_export.columns:
                logger.warning("Field `%s` not present in export", field)
                logger.info(
                    "In Spotlight Preferences > Lightwright, add `%s` to the export fields list.",
                    field,
                )
                self.vw_export[field] = ""

    # Note: Firefox really doesn't like printing 1px borders with border-collapse: collapse
    def default_table_style(self, width: int = 100) -> list[CSSDict]:
        """Return a Style dict with default table styling."""
        return [
            {
                "selector": "",
                "props": "border-spacing: 0px; border-collapse: collapse; "
                "line-height: 1.2; break-inside: auto; width: 100%;",
            },
            {"selector": "tr", "props": "break-after: auto; "},
            {"selector": "td", "props": "padding: 1px; "},
            {
                "selector": "tbody",
                "props": f"display: table; width: {width}%; margin: 0 auto;",
            },
            {
                "selector": "thead tr:not(.generatedMarginals)",
                "props": f"display: table; width: {width}%; margin: 0 auto;",
            },
        ]

    def pagebreak_repeated_index(self, index_name: str) -> list[CSSDict]:
        """Disallow pagebreaks between index fields with the same number.

        Arguments:
            index_name: Field of the dataframe that is considered the index for this paperwork type
                ex. for a channel hookup, this would be "Chan"

        """
        idxs = np.where(self.df[index_name] == self.formatting_quirks.empty_str)

        selector_list = []
        # We want to select the row before a repeated channel,
        # but CSS selectors index from 1. These cancel out.
        selector_list.extend(f"tr:nth-child({i - 1 + 1})" for i in np.transpose(*idxs))

        style_list = []
        style_list.extend({"selector": i, "props": "break-after: avoid;"} for i in selector_list)

        return style_list

    def generate_metadata(self) -> str:
        """Generate HTML metadata from show data."""
        if self.show_data is None:
            return f"""
            <head>
                <meta charset="utf-8">
                <title>{self.display_name}</title>
                <meta name="description" content="{self.display_name}">
                <meta name="generator" content="Lighting Paperwork">
            </head>
            """

        return f"""
        <head>
            <meta charset="utf-8">
            <title>{self.display_name}</title>
            <meta name="description" content="{self.display_name}">
            <meta name="author" content="{self.show_data.ld_name}">
            <meta name="generator" content="Lighting Paperwork">
            <meta name="dcterms.created" content="{self.show_data.print_date()}"
        </head>
        """

    def wrap_table(self, html: str) -> str:
        """Wrap a generated HTML table with bookmarks and anchors."""
        return f"""
        <div id="{self.display_name.replace(" ", "")}" class="report-container"
        style="break-after: page; bookmark-level: 1;
        bookmark-label: '{self.display_name}'; bookmark-state: open;">
            {html}
        </div>
        """

    def generate_header_footer(self, uuid: str) -> tuple[str, str]:
        """Generate a header and footer from show data."""
        if self.show_data is None:
            header_html = self.generate_header(
                uuid, center=StyledContent(self.display_name, f"{self.style.title.to_css()}")
            )
            footer_html = self.generate_footer(
                uuid, left=StyledContent(self.display_name, self.style.marginals.to_css())
            )
        else:
            header_html = self.generate_header(
                uuid,
                right=StyledContent(
                    f"{self.show_data.show_name or ''}<br>{self.show_data.ld_name or ''}",
                    self.style.marginals.to_css() + "margin-bottom: 5%; ",
                ),
                center=StyledContent(self.display_name, f"{self.style.title.to_css()}"),
                left=StyledContent(
                    f"{self.show_data.print_date()}<br>{self.show_data.revision or ''}",
                    self.style.marginals.to_css() + "margin-bottom: 5%; ",
                ),
            )
            footer_html = self.generate_footer(
                uuid,
                left=StyledContent(self.display_name, self.style.marginals.to_css()),
            )

        return (header_html, footer_html)

    @staticmethod
    def generate_header(
        uuid: str,
        left: StyledContent | None = None,
        center: StyledContent | None = None,
        right: StyledContent | None = None,
    ) -> str:
        """Generate the HTML for the table header.

        The generated style :func:`generate_page_style` will hook each `span` into
            a running element for printing to become the page marginal elements.
        The `span`s should semantically be divs probably but those break weasyprint.
        """
        left = StyledContent() if left is None else left
        center = StyledContent() if center is None else center
        right = StyledContent() if right is None else right
        return f"""
        <div id="header_{uuid}" style="display:grid;grid-auto-flow:column;grid-auto-columns:auto;">
            <span class="top-left-{uuid}" id="header_left_{uuid}"
            style="text-align:left;{left.style}">{left.content}</span>
            <span class="top-center-{uuid}" id="header_center_{uuid}"
            style="text-align:center;{center.style}">{center.content}</span>
            <span class="top-right-{uuid}" id="header_right_{uuid}"
            style="text-align:right;{right.style}">{right.content}</span>
        </div>
        """

    @staticmethod
    # pylint: disable-next=too-many-arguments, too-many-positional-arguments
    def generate_footer(
        uuid: str,
        left: StyledContent | None = None,
        center: StyledContent | None = None,
        right: StyledContent | None = None,
    ) -> str:
        """Generate the HTML for the table footer.

        The generated style :func:`generate_page_style` will hook each `span` into
            a running element for printing to become the page marginal elements.
        The `span`s should semantically be divs probably but those break weasyprint.
        """
        left = StyledContent() if left is None else left
        center = StyledContent() if center is None else center
        right = StyledContent() if right is None else right
        return f"""
        <div id="footer_{uuid}" style="display:grid;grid-auto-flow:column;grid-auto-columns:1fr">
            <span class="bottom-left-{uuid}" id="bottom_left_{uuid}"
            style="text-align:left;{left.style}">{left.content}</span>
            <span class="bottom-center-{uuid}" id="bottom_center_{uuid}"
            style="text-align:center;{center.style}">{center.content}</span>
            <span class="bottom-right-{uuid}" id="bottom_right_{uuid}"
            style="text-align:right;{right.style}">{right.content}</span>
        </div>
        """

    @staticmethod
    def generate_page_style(
        uuid: str, pagenum_pos: str | None = None, pagenum_style: str = ""
    ) -> str:
        """Generate a <style> for the table header and footer.

        This establishes the header and footer elements as running, and will insert them
            in the page marginals during printing instead of embedded in the table.
        """
        style = ""
        page_style = ""
        for side in ["left", "center", "right"]:
            for pos in ["top", "bottom"]:
                location_name = f"{pos}-{side}"
                var_name = f"{pos}{side.capitalize()}"

                if pagenum_pos == location_name:
                    page_style += f"""
                        @{location_name} {{
                            content: "Page " counter(page) " of " counter(pages);
                            {pagenum_style}
                        }}
                    """
                else:
                    style += f"""
                        .{location_name}-{uuid} {{
                            position: running({var_name}-{uuid});
                        }}
                        """
                    page_style += f"""
                        @{location_name} {{
                            content: element({var_name}-{uuid});
                        }}
                    """

        return f"""
        <style>
        {style}
        @page {{
            {page_style}
        }}
        </style>
        """
