# type: ignore[reportAttributeAccessIssue]
"""Configure PyTest."""

import pytest

from lighting_paperwork.vectorworks_xml import VWExport


def pytest_configure(config):  # noqa: ARG001
    pytest.NUM_INSTRUMENTS = 41
    pytest.NUM_DUMB_ACCS = 6
    pytest.NUM_SMART_ACCS = 5
    pytest.NUM_POSITIONS = 16


@pytest.fixture
def vwx_export_file():
    return "./tests/TestFile.xml"


@pytest.fixture
def vwx_export(vwx_export_file):
    return VWExport(vwx_export_file).export_df()
