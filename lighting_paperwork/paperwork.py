"""Base paperwork generation class"""

import logging
import re
from abc import ABC, abstractmethod
from typing import List, Optional, Self

import pandas as pd
import openpyxl

from .helpers import FontStyle, ShowData, FormattingQuirks, excel_quirks, html_quirks
from .style import BaseStyle, default_style
import lighting_paperwork.excel_formatter as excel_formatter

logger = logging.getLogger(__name__)


class PaperworkGenerator(ABC):
    """Base paperwork generation class"""

    def __init__(
        self,
        vw_export: pd.DataFrame,
        show_data: Optional[ShowData] = None,
        style: BaseStyle = default_style,
        border_weight: float = 1,
    ) -> None:
        self.vw_export = vw_export
        self.df = self.vw_export.copy()
        self.show_data = show_data
        self.style = style
        # 1px doesn't render right on Firefox, use 1.5px min to workaround.
        self.border_weight = border_weight

    def set_show_data(self, show_name: str, ld_name: str, revision: str) -> None:
        """Save show data for later use"""
        self.show_data = ShowData(
            show_name=show_name, ld_name=ld_name, revision=revision
        )

    display_name: str
    col_widths: list[int]
    page_width: int = 100
    formatting_quirks = html_quirks

    @abstractmethod
    def generate_df(self) -> Self:
        """
        Using `self.vw_export`, generate a DataFrame that contains the
            necessary sorted information for the paperwork type.
        """

    @staticmethod
    @abstractmethod
    def style_data(
        df: pd.DataFrame,
        body_style: FontStyle,
        col_width: List[int],
        border_weight: float,
        quirks: FormattingQuirks,
    ) -> pd.DataFrame:
        """Styles the data for a table"""

    @staticmethod
    @abstractmethod
    def style_fields(
        index: pd.Series,
        header_style: FontStyle,
        col_width: List[int],
        border_weight: float,
    ) -> List[str]:
        """Styles the fields (i.e. headers) for a table"""

    @abstractmethod
    def _make_common(self) -> pd.io.formats.style.Styler:
        """Runs common make tasks for html and excel"""

    @abstractmethod
    def make_html(self) -> str:
        """Generates a formatted HTML table from the generated DataFrame"""

    def make_excel(self, excel_path: str) -> None:
        """Adds a sheet to an Excel file with the formatted DataFrame"""
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
    def verify_width(width: List[int]) -> bool:
        """Verifies that the col widths remain less than 100%"""
        if sum(width) > 100:
            logger.warning("Col widths too long: used %i%%!", sum(width))
            return False

        return True

    def combine_instrtype(self) -> Self:
        """
        Combines the Instrument Type and Power fields into one
        """
        instload = []
        for _, row in self.df.iterrows():
            # Consistently format power
            if row["Wattage"] != "":
                # we want it to be [number]W
                power = re.sub(r"[^\d\.]", "", row["Wattage"])
                powerstr = power + "W"
            else:
                power = None
                powerstr = None

            # Make sure power shows up once, after the instrument type
            if powerstr is None:
                tmp = row["Instrument Type"]
            else:
                # Remove from instrument type (if existing)
                instrtype = re.sub(rf"\s*{power}\w+\s*", "", row["Instrument Type"])
                tmp = instrtype + " " + powerstr

            # If accessory, add that here
            # if row["Accessory Inventory"] != "":
            #    tmp += ", " + row["Accessory Inventory"]

            instload.append(tmp)

        # Clean up by replacing old cols with new one
        # TODO: Get accessories in here
        # df.drop(["Instrument Type", "Wattage", "Accessory Inventory"], axis=1, inplace=True)
        new_df = self.df.drop(["Instrument Type", "Wattage"], axis=1)
        new_df["Instr Type & Load"] = instload

        self.df = new_df
        return self

    def combine_gelgobo(self) -> Self:
        """
        Combines the Gel and Gobo fields into one.
        Only operates on Gobo 1 field, not Gobo 2.
        """
        gelgobo = []
        for _, row in self.df.iterrows():
            # If no gel replace with N/C
            if row["Color"] == "":
                tmp = "N/C"
            else:
                tmp = row["Color"]

            # Append gobo if exists
            if row["Gobo 1"] != "":
                tmp += ", T: " + row["Gobo 1"]

            gelgobo.append(tmp)

        # Clean up by replacing old cols with new one
        new_df = self.df.drop(["Color", "Gobo 1"], axis=1)
        new_df["Color & Gobo"] = gelgobo
        self.df = new_df

        return self

    def format_address_slash(self) -> Self:
        """
        Formats an absolute address into a Universe/Address string.
        """
        for row in self.df.itertuples():
            absaddr = int(self.df.at[row.Index, "Absolute Address"])
            if absaddr == 0:
                # If no address set, replace it with a blank
                self.df.at[row.Index, "Absolute Address"] = (
                    self.formatting_quirks.empty_str
                )
            else:
                universe = int((absaddr - 1) / 512) + 1

                if universe == 1:
                    address = absaddr
                    self.df.at[row.Index, "Absolute Address"] = f"{address}"
                else:
                    address = ((absaddr - 1) % 512) + 1
                    self.df.at[row.Index, "Absolute Address"] = f"{universe}/{address}"

        slashed_df = self.df.rename(columns={"Absolute Address": "Addr"})
        self.df = slashed_df
        return self

    def repeated_channels(self) -> Self:
        """
        Formats repeated channel numbers to use `"` to represent repeated data.
        """
        prev_row = None
        for _, data in self.df.iterrows():
            if prev_row is None:
                prev_row = data
                continue

            if data["Chan"] == prev_row["Chan"]:
                # Repeated channel!
                data["Chan"] = self.formatting_quirks.empty_str
                for idx, val in data.items():
                    if idx == "U#":
                        # Do repeat U# to avoid confusion
                        continue
                    if val == "":
                        # Don't "-ify empty fields
                        continue
                    if data[idx] == prev_row[idx]:
                        data[idx] = '"'
            else:
                prev_row = data

        return self

    def abbreviate_col_names(self) -> Self:
        """
        Abbreviates common column names.
        """
        self.df = self.df.rename(
            columns={"Channel": "Chan", "Unit Number": "U#", "Address": "Addr"}
        )
        return self

    # Note: Firefox really doesn't like printing 1px borders with border-collapse: collapse
    def default_table_style(self, width=100):
        """
        Returns a Style dict with default table styling.
        """
        return [
            {
                "selector": "",
                "props": "border-spacing: 0px; border-collapse: initial; "
                "line-height: 1.2; break-inside: auto; width: 100%;",
            },
            {"selector": "tr", "props": "break-inside: avoid; break-after: auto; "},
            {"selector": "td", "props": "padding: 1px;"},
            {
                "selector": "tbody",
                "props": f"display: table; width: {width}%; margin: 0 auto;",
            },
            {
                "selector": "thead tr:not(.generatedMarginals)",
                "props": f"display: table; width: {width}%; margin: 0 auto;",
            },
        ]

    def generate_metadata(self) -> str:
        """Generates HTML metadata from show data"""
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
        """Wraps a generated HTML table with bookmarks and anchors"""
        return f"""
        <div id="{self.display_name.replace(" ", "")}" class="report-container" style="break-after: page; bookmark-level: 1; bookmark-label: '{self.display_name}'; bookmark-state: open;">
            {html}
        </div>
        """

    def generate_header_footer(self, uuid: str) -> tuple[str, str]:
        """Generates a header and footer from show data"""
        if self.show_data is None:
            header_html = self.generate_header(
                uuid,
                content_center=self.display_name,
                style_center=f"{self.style.title.to_css()}",
            )
            footer_html = self.generate_footer(
                uuid,
                style_left=self.style.marginals.to_css(),
                content_left=self.display_name,
            )
        else:
            header_html = self.generate_header(
                uuid,
                content_right=f"{self.show_data.show_name}<br>{self.show_data.ld_name}",
                content_left=f"{self.show_data.print_date()}<br>{self.show_data.revision}",
                content_center=self.display_name,
                style_right=self.style.marginals.to_css() + "margin-bottom: 5%; ",
                style_left=self.style.marginals.to_css() + "margin-bottom: 5%; ",
                style_center=f"{self.style.title.to_css()}",
            )
            footer_html = self.generate_footer(
                uuid,
                style_left=self.style.marginals.to_css(),
                content_left=self.display_name,
            )

        return (header_html, footer_html)

    @staticmethod
    # pylint: disable-next=too-many-arguments, too-many-positional-arguments
    def generate_header(
        uuid: str,
        content_left: str = "",
        content_center: str = "",
        content_right: str = "",
        style_left: str = "",
        style_center: str = "",
        style_right: str = "",
    ) -> str:
        """
        Generates the HTML for the table header.
        The generated style :func:`generate_page_style` will hook each `span` into
            a running element for printing to become the page marginal elements.
        The `span`s should semantically be divs probably but those break weasyprint.
        """
        return f"""
        <div id="header_{uuid}" style="display:grid;grid-auto-flow:column;grid-auto-columns:auto;">
            <span class="top-left-{uuid}" id="header_left_{uuid}"
            style="text-align:left;{style_left}">{content_left}</span>
            <span class="top-center-{uuid}" id="header_center_{uuid}"
            style="text-align:center;{style_center}">{content_center}</span>
            <span class="top-right-{uuid}" id="header_right_{uuid}"
            style="text-align:right;{style_right}">{content_right}</span>
        </div>
        """

    @staticmethod
    # pylint: disable-next=too-many-arguments, too-many-positional-arguments
    def generate_footer(
        uuid: str,
        content_left: str = "",
        content_center: str = "",
        content_right: str = "",
        style_left: str = "",
        style_center: str = "",
        style_right: str = "",
    ) -> str:
        """
        Generates the HTML for the table footer.
        The generated style :func:`generate_page_style` will hook each `span` into
            a running element for printing to become the page marginal elements.
        The `span`s should semantically be divs probably but those break weasyprint.
        """
        return f"""
        <div id="footer_{uuid}" style="display:grid;grid-auto-flow:column;grid-auto-columns:1fr">
            <span class="bottom-left-{uuid}" id="bottom_left_{uuid}"
            style="text-align:left;{style_left}">{content_left}</span>
            <span class="bottom-center-{uuid}" id="bottom_center_{uuid}"
            style="text-align:center;{style_center}">{content_center}</span>
            <span class="bottom-right-{uuid}" id="bottom_right_{uuid}"
            style="text-align:right;{style_right}">{content_right}</span>
        </div>
        """

    @staticmethod
    def generate_page_style(
        uuid: str, pagenum_pos: Optional[str] = None, pagenum_style: str = ""
    ) -> str:
        """
        Generates a <style> for the table header and footer.
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

        html = f"""
        <style>
        {style}
        @page {{
            {page_style}
        }}
        </style>
        """

        return html
