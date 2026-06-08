# SEO-Analysis

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

A Python script to gain some insights from a domain and a list of keywords.

## Description

SEO-Analysis reads a list of domains and keywords from an Excel workbook, fetches each domain, checks the page content against the keywords, and writes the insights back into the workbook with colour-coded cells (red, yellow, white) indicating how well each keyword is represented.

## Requirements

- Python 3
- Dependencies: `requests`, `beautifulsoup4`, `openpyxl`

Install them with:

```
pip install -r requirements.txt
```

## Usage

1. Open `Test.xlsx`.
2. Add your keywords and domains, one pair per row. The same domain may appear on multiple rows with different keywords.
3. Save `Test.xlsx`.
4. Run the script:

```
python3 SEO.py
```

5. The workbook is updated in place with the analysis and colour-coded cells.

You may also find the companion project useful: https://github.com/com-puter-tips/Links-Extractor

## Citation

If you use this software, please cite it using the metadata in [CITATION.cff](CITATION.cff).

## License

Distributed under the GNU General Public License v3.0. See [LICENSE](LICENSE).
