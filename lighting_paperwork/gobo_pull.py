"""Generator for a gobo pull list"""

import logging
from typing import List, Self

import numpy as np
import pandas as pd
from pandas.io.formats.style import Styler

from .helpers import FontStyle, FormattingQuirks
from .paperwork import PaperworkGenerator

logger = logging.getLogger(__name__)


class GoboPullList(PaperworkGenerator):
    """
    Generates a gobo pull list with gobo name and quantity.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    col_widths = [80, 20]
    page_width = 40
    display_name = "Gobo Pull List"

    def generate_df(self) -> Self:
        filter_fields = ["Gobo 1", "Gobo 2"]
        chan_fields = pd.DataFrame(self.vw_export[filter_fields], columns=filter_fields)
        gobo_list = []

        for _, row in chan_fields.iterrows():
            if row["Gobo 1"].strip() != "":
                gobo_list.append(row["Gobo 1"])
            if row["Gobo 2"].strip() != "":
                gobo_list.append(row["Gobo 2"])

        gobo_name, gobo_count = np.unique(gobo_list, return_counts=True)

        self.df = pd.DataFrame(zip(gobo_name, gobo_count), columns=["Gobo Name", "Qty"])

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

        style_df = style_df.astype(str)
        for index, _ in df.iterrows():
            style_df.loc[index, :] = ""
            style_df.loc[index, :] += f"border-bottom: {border_style}; "

        # Set font based on column
        for col_name, _ in style_df.items():
            width_idx = style_df.columns.get_loc(col_name)
            style_df[col_name] += (
                f"{body_style.to_css()}; vertical-align: middle; width: {col_width[width_idx]}%; "
            )

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
            if val in ["Gobo Name"]:
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

        logger.info("Generated gobo pull list.")

        html = styled.to_html(
            generated_header=header_html,
            generated_footer=footer_html,
            generated_page_style=page_style,
        )
        html = self.wrap_table(html)

        return html
