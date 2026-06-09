"""Backward-compatible entry point.

The analysis logic now lives in the ``seo_analysis`` package under ``src/``.
This shim preserves the original usage: from the repo root, run

    python SEO.py

and it analyses ``Test.xlsx`` (Sheet1) in the current directory, exactly as
before -- no installation required.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from seo_analysis import analyze

if __name__ == "__main__":
    analyze("Test.xlsx", "Sheet1")
