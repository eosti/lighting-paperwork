import logging
import re
from abc import ABC, abstractmethod
from typing import List, Optional, Self

import pandas as pd

from helpers import FontStyle, ShowData
from style import BaseStyle, DefaultStyle

logger = logging.getLogger(__name__)


class PaperworkGenerator(ABC):
    def __init__(
        self,
        vw_export: pd.DataFrame,
        show_data: Optional[ShowData] = None,
        style: BaseStyle = DefaultStyle,
        border_weight: float = 1,
    ) -> None:
        self.vw_export = vw_export
        self.df = self.vw_export.copy()
        self.show_data = show_data
        self.style = style
        # 1px doesn't render right on Firefox, use 1.5px min to workaround.
        self.border_weight = border_weight

    def set_show_data(self, show_name: str, ld_name: str, revision: str) -> None:
        self.show_data = ShowData(
            show_name=show_name, ld_name=ld_name, revision=revision
        )

    @abstractmethod
    def generate_df(self) -> pd.DataFrame:
        pass

    @staticmethod
    @abstractmethod
    def style_data(df: pd.DataFrame, body_style: FontStyle) -> pd.DataFrame:
        pass

    @staticmethod
    @abstractmethod
    def style_fields(
        index: pd.Series, header_style: FontStyle, col_width: List[int]
    ) -> List[str]:
        pass

    @staticmethod
    def verify_width(width: List[int]) -> bool:
        if sum(width) > 100:
            logger.warn(f"Col widths too long: used {sum(width)}%!")
            return False

        return True

    def combine_instrtype(self) -> Self:
        """
        Combines the Instrument Type and Power fields into one
        """
        instload = []
        for index, row in self.df.iterrows():
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
        # Only operates on Gobo 1
        gelgobo = []
        for index, row in self.df.iterrows():
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
        for row in self.df.itertuples():
            absaddr = int(self.df.at[row.Index, "Absolute Address"])
            if absaddr == 0:
                # If no address set, replace it with a blank
                self.df.at[row.Index, "Absolute Address"] = ""
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
        prev_row = None
        for index, data in self.df.iterrows():
            if prev_row is None:
                prev_row = data
                continue

            if data["Chan"] == prev_row["Chan"]:
                # Repeated channel!
                data["Chan"] = "&nbsp;"
                for idx, val in data.items():
                    if idx == "U#":
                        # Do repeat U# to avoid confusion
                        continue
                    elif val == "":
                        # Don't "-ify empty fields
                        continue
                    elif data[idx] == prev_row[idx]:
                        data[idx] = '"'
            else:
                prev_row = data

        return self

    def abbreviate_col_names(self) -> Self:
        self.df = self.df.rename(
            columns={"Channel": "Chan", "Unit Number": "U#", "Address": "Addr"}
        )
        return self

    # Note: Firefox really doesn't like printing 1px borders with border-collapse: collapse
    default_table_style = [
        {
            "selector": "",
            "props": "border-spacing: 0px; border-collapse: initial; line-height: 1.2; break-inside: auto; width: 100%;",
        },
        {"selector": "tr", "props": "break-inside: avoid; break-after: auto; "},
        {"selector": "td", "props": "padding: 1px;"},
    ]

    @staticmethod
    def generate_header(
        uuid: str,
        content_left: str = "",
        content_center: str = "",
        content_right: str = "",
        style_left: str = "",
        style_center: str = "",
        style_right: str = "",
    ) -> str:
        return f"""
        <div id="header_{uuid}" style="display:grid;grid-auto-flow:column;grid-auto-columns:1fr">
            <div class="top-left" id="header_left_{uuid}" style="text-align:left;{style_left}">{content_left}</div>
            <div class="top-center" id="header_center_{uuid}" style="text-align:center;{style_center}">{content_center}</div>
            <div class="top-right" id="header_right_{uuid}" style="text-align:right;{style_right}">{content_right}</div>
        </div>
        """

    @staticmethod
    def generate_footer(
        uuid: str,
        content_left: str = "",
        content_center: str = "",
        content_right: str = "",
        style_left: str = "",
        style_center: str = "",
        style_right: str = "",
    ) -> str:
        return f"""
        <div id="footer_{uuid}" style="display:grid;grid-auto-flow:column;grid-auto-columns:1fr">
            <div class="bottom-left" id="bottom_left_{uuid}" style="text-align:left;{style_left}">{content_left}</div>
            <div class="bottom-center" id="bottom_center_{uuid}" style="text-align:center;{style_center}">{content_center}</div>
            <div class="bottom-right" id="bottom_right_{uuid}" style="text-align:right;{style_right}">{content_right}</div>
        </div>
        """

    @staticmethod
    def generate_page_style(
        pagenum_pos: Optional[str] = None, pagenum_style: str = ""
    ) -> str:
        html = """
        <style>
        @page {
        """
        for side in ["left", "center", "right"]:
            for pos in ["top", "bottom"]:
                class_name = f"{pos}-{side}"
                var_name = f"{pos}{side.capitalize()}"

                if pagenum_pos == class_name:
                    html += f"""
                        @{class_name} {{
                            content: "Page " counter(page) " of " counter(pages);
                            {pagenum_style}
                        }}
                    """
                else:
                    html += f"""
                        @{class_name} {{
                            content: element({var_name});
                        }}
                        .{class_name} {{
                            position: running({var_name})
                        }}
                """

        html += """
        }
        </style>
        """

        return html
