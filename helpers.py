from dataclasses import dataclass
import re


@dataclass
class ShowData:
    show_name: str
    ld_name: str
    revision: str


@dataclass
class Gel:
    name: str
    name_sort: str
    company: str


def parse_gel(gel: str) -> Gel:
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

    return Gel(gel, gelsort, company)
