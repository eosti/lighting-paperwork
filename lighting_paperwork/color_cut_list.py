"""Generator for a color cut list"""

import logging
from typing import List, Self

import pandas as pd
from natsort import natsort_keygen
from pandas.io.formats.style import Styler

from .helpers import FontStyle, Gel, FormattingQuirks
from .paperwork import PaperworkGenerator

logger = logging.getLogger(__name__)


class ColorCutList(PaperworkGenerator):
    """
    Generates a color pull list with color, size, and quantity.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    col_widths = [34, 43, 23]
    page_width = 30
    display_name = "Color Cut List"

    def generate_df(self) -> Self:
        filter_fields = ["Color", "Frame Size"]

        self.df = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)
        # Seperate colors and diffusion into dict list
        color_dict = []
        for _, row in self.df.iterrows():
            if row["Frame Size"] != "" and not pd.isnull(row["Frame Size"]):
                framesize = row["Frame Size"]
            else:
                framesize = "Unknown"
            if (
                row["Color"].strip() != ""
                and row["Color"] != "N/C"
                and not pd.isnull(row["Color"])
            ):
                for i in row["Color"].strip().split("+"):
                    # Works for single gels too
                    if len(i.split("x")) > 1:
                        # Repeat gel situation (i.e. L201x2)
                        for _ in range(0, int(i.split("x")[1])):
                            gel = Gel.parse_name(i.split("x")[0])
                            color_dict.append(
                                {
                                    "Color": gel.name,
                                    "Frame Size": framesize,
                                    "Company": gel.company,
                                    "Sort": gel.name_sort,
                                }
                            )
                    else:
                        # Normal single gel
                        gel = Gel.parse_name(i)
                        color_dict.append(
                            {
                                "Color": gel.name,
                                "Frame Size": framesize,
                                "Company": gel.company,
                                "Sort": gel.name_sort,
                            }
                        )

        colors = pd.DataFrame.from_dict(color_dict)
        colors = (
            colors.groupby(["Color", "Frame Size", "Sort"])["Color"]
            .count()
            .reset_index(name="Count")
        )
        # Hack for that silly Rosco company: 3xx values become xx.3
        colors = colors.sort_values(by=["Sort", "Frame Size"], key=natsort_keygen())
        colors = colors.drop(["Sort"], axis=1)

        self.df = colors
        return self

    @staticmethod
    def style_data(
        df: pd.DataFrame,
        body_style: FontStyle,
        col_width: List[int],
        border_weight: float,
        quirks: FormattingQuirks,
    ):
        border_style = f"{border_weight}px solid black"
        style_df = df.copy()
        # Set borders based on color data
        prev_row = (None, None)
        style_df = style_df.astype(str)
        for index, data in df.iterrows():
            style_df.loc[index, :] = ""
            if prev_row == (None, None):
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
                prev_row = (index, data)
                continue

            if data["Color"] == prev_row[1]["Color"]:
                style_df.loc[prev_row[0], "Color"] += "border-bottom: none; "
                style_df.loc[index, "Color"] += f"color: {quirks.hidden_fmt}; "
                style_df.loc[index, :] += f"border-bottom: {border_style}; "
            else:
                style_df.loc[index, :] += f"border-bottom: {border_style}; "

            prev_row = (index, data)

        # Set font based on column
        for col_name, _ in style_df.items():
            width_idx = style_df.columns.get_loc(col_name)
            style_df[col_name] += (
                f"{body_style.to_css()}; vertical-align: middle; width: {col_width[width_idx]}%; "
            )

            if col_name in ["Color"]:
                style_df[col_name] += "text-align: center; "
            else:
                style_df[col_name] += "text-align: left; "

        return style_df

    @staticmethod
    def style_fields(
        index: pd.Series,
        header_style: FontStyle,
        col_width: List[int],
        border_weight: float,
    ) -> List[str]:
        PaperworkGenerator.verify_width(col_width)
        style = [
            f"{header_style.to_css()}; border-bottom: {border_weight}px solid black; "
            for _ in index
        ]

        for idx, val in enumerate(index):
            if val in ["Color"]:
                style[idx] += "text-align: center; "
            else:
                style[idx] += "text-align: left; "

        for idx, _ in enumerate(index):
            style[idx] += f"width: {col_width[idx]}%; "

        return style

    def _make_common(self) -> pd.io.formats.style.Styler:
        self.generate_df()

        styled = Styler.from_custom_template(".", "header_footer.tpl")(self.df)
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
            type(self).style_fields,
            header_style=self.style.field,
            col_width=self.col_widths,
            border_weight=self.border_weight,
            axis=1,
        )

        return styled

    def make_html(self) -> str:
        styled = self._make_common()

        styled = styled.set_table_attributes('class="paperwork-table"')
        styled = styled.set_table_styles(
            self.default_table_style(width=self.page_width), overwrite=False
        )

        header_html, footer_html = self.generate_header_footer(styled.uuid)
        page_style = self.generate_page_style(
            styled.uuid, "bottom-right", self.style.marginals.to_css()
        )

        logger.info("Generated color cut list.")

        html = styled.to_html(
            generated_header=header_html,
            generated_footer=footer_html,
            generated_page_style=page_style,
        )
        html = self.wrap_table(html)

        return html
