"""Sphinx configuration for the docxium documentation.

The documentation has two parts: the docxium pages in this directory and the
python-docx documentation under ``python-docx/``.
"""

from pathlib import Path

import docx
import docxium

DOCS = Path(__file__).parent

project = "docxium"
author = "CMOS"
copyright = "2026, CMOS"
version = release = docxium.__version__

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.intersphinx",
    "sphinx.ext.viewcode",
]

# The python-docx pages use substitutions such as |Document| and |docx-version|.
rst_epilog = (DOCS / "python-docx" / "substitutions.inc").read_text(
    encoding="utf-8"
) + f"\n.. |docx-version| replace:: {docx.__version__}\n"

exclude_patterns = ["_build"]
pygments_style = "sphinx"
intersphinx_mapping = {"python": ("https://docs.python.org/3", None)}

html_theme = "shibuya"
html_static_path = ["_static"]
html_theme_options = {
    "accent_color": "blue",
    "github_url": "https://github.com/cmostw/docxium",
}
