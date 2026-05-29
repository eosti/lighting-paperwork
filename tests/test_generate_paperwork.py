"""Tests for the generate_paperwork CLI."""

from unittest.mock import patch

from lighting_paperwork.generate_paperwork import main


def test_smoke_test():
    """Basic smoke test.

    Does the program run at all?
    """
    testargs = ["testing.py", "tests/TestFile.xml"]
    with patch("sys.argv", testargs):
        main()
