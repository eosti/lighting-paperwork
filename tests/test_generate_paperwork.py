"""Tests for the generate_paperwork CLI."""

from lighting_paperwork.generate_paperwork import main


def test_smoke_test():
    """Basic smoke test.

    Does the program run at all?
    """
    main(
        [
            "tests/TestFile.xml",
        ]
    )
