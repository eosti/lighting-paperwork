"""Tests for the generate_paperwork CLI."""

from unittest.mock import patch

import pytest
from pydantic import ValidationError

from lighting_paperwork.generate_paperwork import main


def test_smoke_test():
    """Basic smoke test.

    Does the program run at all?
    """
    testargs = ["testing.py", "tests/TestFile.xml"]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0


def test_yaml_input(tmp_path):
    """Test if YAML parsing isn't broken."""
    settings = """
log_level: DEBUG
data_file: "tests/TestFile.xml"
paperwork:
  show_info:
    show_name: "Showy the Showsicle"
    ld_name: "eosti"
    revision: "Rev. Z"
    """
    settings_file = tmp_path / "settings.yaml"
    settings_file.write_text(settings)

    testargs = ["testing.py", str(settings_file)]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0


def test_invalid_yaml_nonexistent_data(tmp_path):
    """Test for non-existent data file."""
    settings = """
log_level: DEBUG
data_file: "tests/nothere.xml"
paperwork:
  show_info:
    show_name: "Showy the Showsicle"
    ld_name: "eosti"
    revision: "Rev. Z"
    """
    settings_file = tmp_path / "settings.yaml"
    settings_file.write_text(settings)

    testargs = ["testing.py", str(settings_file)]
    with patch("sys.argv", testargs), pytest.raises(ValidationError):
        main()


def test_nonexistent_yaml():
    """Test for non-existent YAML."""
    testargs = ["testing.py", "notaconfig.yaml"]
    with patch("sys.argv", testargs), pytest.raises(ValidationError):
        main()
