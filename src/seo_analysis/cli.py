"""Command-line interface for SEO-Analysis."""

import argparse

from . import __version__
from .analyzer import analyze


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="seo-analysis",
        description=(
            "Spreadsheet-driven on-page SEO analysis. Reads keywords (column A) "
            "and URLs (column B) from an .xlsx workbook, fetches each page, and "
            "fills columns C-T with title, description, header, image, link, "
            "video and list-item analysis, colour-coding the cells."
        ),
    )
    parser.add_argument(
        "file",
        nargs="?",
        default="Test.xlsx",
        help="Path to the .xlsx workbook (default: Test.xlsx).",
    )
    parser.add_argument(
        "--sheet",
        default="Sheet1",
        help="Worksheet name to read and write (default: Sheet1).",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=None,
        help="Per-request timeout in seconds (default: no timeout).",
    )
    parser.add_argument(
        "--user-agent",
        default=None,
        help="Custom User-Agent header for requests (default: requests default).",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s " + __version__,
    )

    args = parser.parse_args(argv)
    analyze(
        filepath=args.file,
        sheet_name=args.sheet,
        timeout=args.timeout,
        user_agent=args.user_agent,
    )


if __name__ == "__main__":
    main()
