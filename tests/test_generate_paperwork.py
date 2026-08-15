"""Tests for the generate_paperwork CLI."""

from unittest.mock import patch

import pytest
from pydantic import ValidationError

from lighting_paperwork.generate_paperwork import main


def test_xml_input():
    """Basic smoke test.

    Does the program run at all?
    """
    testargs = ["testing.py", "tests/TestFile.xml"]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0


def test_csv_input():
    """Basic smoke test.

    Does the program run at all?
    """
    testargs = ["testing.py", "tests/TestFile.csv"]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0


def test_yaml_input(tmp_path):
    """Test if YAML parsing isn't broken."""
    settings = f"""
log_level: DEBUG
data_file: "tests/TestFile.xml"
output_dir: {tmp_path}
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
    settings = f"""
log_level: DEBUG
data_file: "tests/nothere.xml"
output_dir: {tmp_path}
paperwork:
  show_info:
    show_name: "Showy the Showsicle"
    ld_name: "eosti"
    revision: "Rev. Z"
    """
    settings_file = tmp_path / "settings_invalid_yaml.yaml"
    settings_file.write_text(settings)

    testargs = ["testing.py", str(settings_file)]
    with patch("sys.argv", testargs), pytest.raises(ValidationError):
        main()


def test_nonexistent_yaml():
    testargs = ["testing.py", "notaconfig.yaml"]
    with patch("sys.argv", testargs), pytest.raises(ValidationError):
        main()


def test_pdf_output(tmp_path):
    settings = f"""
log_level: DEBUG
data_file: "tests/TestFile.xml"
output_type: pdf
output_dir: {tmp_path}
    """
    settings_file = tmp_path / "settings_pdf.yaml"
    settings_file.write_text(settings)

    testargs = ["testing.py", str(settings_file)]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0

    output_file = tmp_path / "Paperwork.pdf"
    assert output_file.exists()


def test_html_output(tmp_path):
    settings = f"""
log_level: DEBUG
data_file: "tests/TestFile.xml"
output_type: html
output_dir: {tmp_path}
    """
    settings_file = tmp_path / "settings_html.yaml"
    settings_file.write_text(settings)

    testargs = ["testing.py", str(settings_file)]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0

    output_file = tmp_path / "Paperwork.html"
    assert output_file.exists()


def test_excel_output(tmp_path):
    settings = f"""
log_level: DEBUG
data_file: "tests/TestFile.xml"
output_type: excel
output_dir: {tmp_path}
    """
    settings_file = tmp_path / "settings_excel.yaml"
    settings_file.write_text(settings)

    testargs = ["testing.py", str(settings_file)]
    with patch("sys.argv", testargs):
        with pytest.raises(SystemExit) as e:
            main()

        assert e.value.code == 0

    output_file = tmp_path / "Paperwork.xlsx"
    assert output_file.exists()
