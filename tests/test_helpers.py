"""Tests for all the helper functions and classes."""

import re

import pytest

from lighting_paperwork.helpers import (
    DMXAddress,
    Gel,
    InstrumentPower,
    parse_frame_size,
)


@pytest.mark.parametrize(
    ("gel_name", "company"),
    [
        ("L202", "Lee"),
        ("R26", "Rosco"),
        ("AP2340", "Apollo"),
        ("G990", "GAM"),
        ("Red", ""),
        ("Lavender", ""),
        ("Gold", ""),
        ("Auburn", ""),
    ],
)
def test_gel_basic(gel_name, company):
    g = Gel._parse_name(gel_name)
    assert isinstance(g, Gel)
    assert g.name == gel_name
    assert g.name_sort == gel_name
    assert g.company == company


def test_gel_special():
    g = Gel._parse_name("R305")
    assert isinstance(g, Gel)
    assert g.name == "R305"
    assert g.name_sort == "R05.3"
    assert g.company == "Rosco"

    g = Gel._parse_name("r05")
    assert isinstance(g, Gel)
    assert g.name == "R05"
    assert g.name_sort == "R05"
    assert g.company == "Rosco"

    g = Gel._parse_name("")
    assert isinstance(g, Gel)
    assert g.name == ""
    assert g.name_sort == ""
    assert g.company == ""

    g = Gel.parse_gel("")
    assert len(g) == 1
    assert g[0].name == ""
    assert g[0].name_sort == ""
    assert g[0].company == ""

    g = Gel.parse_gel("TBD")
    assert len(g) == 1
    assert g[0].name == "TBD"
    assert g[0].name_sort == "TBD"
    assert g[0].company == ""

    g = Gel.parse_gel("Yellow-ish 384")
    assert len(g) == 1
    assert g[0].name == "Yellow-ish 384"
    assert g[0].name_sort == "Yellow-ish 384"
    assert g[0].company == ""

    g = Gel.parse_gel("N/C")
    assert len(g) == 1
    assert g[0].name == "N/C"
    assert g[0].name_sort == "N/C"
    assert g[0].company == ""


def test_gel_complex():
    gels = Gel.parse_gel("L202x2")
    assert len(gels) == 2
    assert gels[0] == Gel._parse_name("L202")
    assert gels[0] == gels[1]

    gels = Gel.parse_gel("R364x5")
    assert len(gels) == 5
    assert gels[0] == Gel._parse_name("R364")
    assert gels[0] == gels[1]
    assert gels[0] == gels[2]
    assert gels[0] == gels[3]
    assert gels[0] == gels[4]

    gels = Gel.parse_gel("R02 + R119")
    assert len(gels) == 2
    assert gels[0] == Gel._parse_name("R02")
    assert gels[1] == Gel._parse_name("R119")

    gels = Gel.parse_gel("R02 + R119 + L202x2")
    assert len(gels) == 4
    assert gels[0] == Gel._parse_name("R02")
    assert gels[1] == Gel._parse_name("R119")
    assert gels[2] == Gel._parse_name("L202")
    assert gels[3] == Gel._parse_name("L202")

    gels = Gel.parse_gel("R02+R119+L202x2")
    assert len(gels) == 4
    assert gels[0] == Gel._parse_name("R02")
    assert gels[1] == Gel._parse_name("R119")
    assert gels[2] == Gel._parse_name("L202")
    assert gels[3] == Gel._parse_name("L202")


@pytest.mark.parametrize(
    ("frame_size", "output_str"),
    [
        ('6.5"', '6.5"'),
        ("6.5", '6.5"'),
        ("6.5in", '6.5"'),
        ("6.5 in", '6.5"'),
        ('16"x24"', '16"x24"'),
        ("16x24", '16"x24"'),
        ("16X24", '16"x24"'),
    ],
)
def test_parse_frame_size(frame_size, output_str):
    assert parse_frame_size(frame_size) == output_str


@pytest.mark.parametrize(
    ("input_addr", "absolute_address"),
    [
        ("4", 4),
        (4, 4),
        (515, 515),
        ("4/17", 1553),
        ("4:17", 1553),
        ("1/123", 123),
        ("1:123", 123),
    ],
)
def test_absolute_address(input_addr, absolute_address):
    assert DMXAddress(input_addr).absolute_address == absolute_address


@pytest.mark.parametrize(
    "input_addr",
    [
        0,
        "0",
        "0/6",
        "0/0",
        "1/513",
        "0:6",
        "0:0",
        "-1/4",
        "6/0",
        "asdf",
        "6/asdf",
        "asdf/6",
        "fourteen",
    ],
)
def test_invalid_address(input_addr):
    with pytest.raises(ValueError):
        DMXAddress(input_addr)


@pytest.mark.parametrize(
    ("input_addr", "output_str"),
    [
        ("4", "1/4"),
        (4, "1/4"),
        (515, "2/3"),
        ("4/17", "4/17"),
        ("4:17", "4/17"),
        (1553, "4/17"),
    ],
)
def test_slash_address(input_addr, output_str):
    assert DMXAddress(input_addr).format_slash() == output_str


@pytest.mark.parametrize(
    ("input_addr", "output_str"),
    [
        ("4", "1:4"),
        (4, "1:4"),
        (515, "2:3"),
        ("4/17", "4:17"),
        ("4:17", "4:17"),
        (1553, "4:17"),
    ],
)
def test_colon_address(input_addr, output_str):
    assert DMXAddress(input_addr).format_colon() == output_str


@pytest.mark.parametrize(
    ("input_addr", "output_str"),
    [
        ("4", "4"),
        (4, "4"),
        (515, "2/3"),
        ("4/17", "4/17"),
        ("4:17", "4/17"),
        (1553, "4/17"),
    ],
)
def test_slash_conditional_address(input_addr, output_str):
    assert DMXAddress(input_addr).format_slash_conditional() == output_str


@pytest.mark.parametrize(
    ("input_power", "output_str"),
    [
        (10, "10W"),
        ("10", "10W"),
        (10.0, "10W"),
        ("10.0", "10W"),
        ("10W", "10W"),
        ("10 W", "10W"),
        ("10kW", "10kW"),
        ("10 kW", "10kW"),
        ("10000", "10kW"),
        (10000, "10kW"),
        (999, "999W"),
        (1000, "1kW"),
        (None, "0W"),
        ("", "0W"),
        ("0kW", "0W"),
    ],
)
def test_power_formatting(input_power, output_str):
    assert InstrumentPower(input_power).format() == output_str


def test_power_formatting_assertion():
    with pytest.raises(ValueError):
        InstrumentPower(-1)
    with pytest.raises(ValueError):
        InstrumentPower(-1.0)
    with pytest.raises(ValueError):
        InstrumentPower("-1")
    with pytest.raises(ValueError):
        InstrumentPower("-1kW")


@pytest.mark.parametrize(
    ("input_str", "output_str"),
    [
        ("INSTR 250W", "250W"),
        ("INSTR 250 W", "250 W"),
        ("INSTR 250", ""),
        ("INSTR 150 250W", "250W"),
        ("INSTR 250.4 W", "250.4 W"),
        ("INSTR 250 kW", "250 kW"),
        ("INSTR 250 W MORE INSTR", "250 W"),
    ],
)
def test_power_regex(input_str, output_str):
    ans = re.search(InstrumentPower.POWER_REGEX, input_str)

    if ans is None:
        assert output_str == ""
    else:
        assert ans.group() == output_str
