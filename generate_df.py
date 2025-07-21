import re
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
from natsort import natsort_keygen, natsorted

from helpers import parse_gel


def gobo_pull(vw_export: pd.DataFrame) -> pd.DataFrame:
    filter_fields = ["Gobo 1", "Gobo 2"]
    chan_fields = pd.DataFrame(vw_export[filter_fields], columns=filter_fields)
    gobo_list = []

    for index, row in chan_fields.iterrows():
        if row["Gobo 1"].strip() != "":
            gobo_list.append(row["Gobo 1"])
        if row["Gobo 2"].strip() != "":
            gobo_list.append(row["Gobo 2"])

    gobo_name, gobo_count = np.unique(gobo_list, return_counts=True)

    return pd.DataFrame(zip(gobo_name, gobo_count), columns=["Gobo Name", "Quantity"])
