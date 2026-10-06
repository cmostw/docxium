<a id="readme-top"></a>

<!-- PROJECT HEADER -->
<br />
<div align="center">
  <h3 align="center">docxium</h3>

  <p align="center">
    python-docx extended with batch operations implemented in Rust.
  </p>

[![MIT License][license-shield]][license-url] [![Rust][Rust-badge]][Rust-url] [![Python][Python-badge]][Python-url]

  <p align="center">
    <a href="https://cmostw.github.io/docxium/"><strong>Documentation »</strong></a>
    &middot;
    <a href="#benchmarks">Benchmarks</a>
    &middot;
    <a href="https://github.com/cmostw/docxium/issues/new?labels=bug">Report Bug</a>
    &middot;
    <a href="https://github.com/cmostw/docxium/issues/new?labels=enhancement">Request Feature</a>
  </p>
</div>

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#about-the-project">About The Project</a></li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#benchmarks">Benchmarks</a></li>
    <li><a href="#license">License</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->
## About The Project

[python-docx](https://github.com/python-openxml/python-docx) adds content to a document one call
at a time. With 50,000 paragraphs of three runs each, building the document takes 39.5 s and
reading all paragraph text back takes 3.5 s (see [Benchmarks](#benchmarks)).

docxium keeps python-docx as the document model and adds batch methods. A batch method receives a
whole list of paragraphs or table rows in one call. Rust generates the WordprocessingML for the
list, and python-docx's own parser inserts it into the document, so the result is an ordinary
python-docx `Document`.

* `import docxium as docx` exposes the batch methods. Names and submodules that docxium does not
  define (`docxium.shared`, `docxium.enum.text`, ...) are forwarded to python-docx.
* Methods that python-docx already provides are python-docx's and are not changed.
* Opening and saving are performed by python-docx and are not changed.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- USAGE EXAMPLES -->
## Usage

Everything python-docx provides is used as python-docx documents it. Only the additions below
are specific to docxium. They are available on the document returned by `docxium.Document()`,
or on an existing `docx.Document` after `docxium.accelerate(doc)`.

```python
import docxium as docx

doc = docx.Document()
doc.add_headings(["Introduction", "Methods"], level=1, alignment="center")
doc.add_paragraphs(
    [
        "plain",
        ["mixed ", ("bold", True, None), ("red", {"size": 18, "color": "FF0000"})],
    ],
    alignment="justify",
)
doc.add_table_from_rows(rows, header=True, col_widths=[docx.shared.Inches(1)] * 3)
texts = doc.extract_paragraph_texts()
```

| method | behavior |
|---|---|
| `add_paragraphs(paragraphs, style=None, alignment=None)` | Appends one paragraph per item. `style` is a style name or object; `alignment` is a `WD_ALIGN_PARAGRAPH` member or its name; both apply to every item. Equivalent to `add_paragraph` and `add_run` calls. |
| `add_headings(texts, level=1, alignment=None)` | Same as `add_paragraphs` with the `Title` style for level 0 and `Heading N` for levels 1-9, as `add_heading` selects. |
| `add_table_from_rows(rows, style=None, *, header=False, col_widths=None)` | Appends a table and returns it. Cells that are not `str` are converted with `str()`. `header=True` writes the first row in bold. `col_widths` holds one width per column (a `docx.shared.Length` or an EMU integer); without it the text width is divided evenly, as `add_table` does. |
| `extract_paragraph_texts()` | Returns the text of every paragraph in document order, including paragraphs inside tables. `Document.paragraphs` lists body-level paragraphs only. |

A paragraph is a `str` or an iterable of runs. A run is a `str`, `(text, bold, italic)`, or
`(text, {...})` where the dict may contain:

| key | value |
|---|---|
| `bold`, `italic`, `underline`, `strike`, `all_caps`, `small_caps` | bool |
| `superscript`, `subscript` | bool (not both) |
| `size` | points, e.g. `10.5` |
| `color` | `"RRGGBB"` (a leading `#` is accepted) |
| `font` | font name |
| `highlight` | color name (`"yellow"`, `"darkBlue"`, `"dark_blue"`) or a `WD_COLOR_INDEX` member |

Invalid input raises `TypeError` or `ValueError`.

The written XML equals python-docx's for the equivalent calls, with one difference: an empty table
cell is written as `<w:p/>`, while python-docx writes `<w:p><w:r/></w:p>`.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- BENCHMARKS -->
## Benchmarks

Speed-up is the python-docx time divided by the docxium time. Each workload runs on one thread;
the value reported is the median of 7 rounds.

* **Workloads:** paragraphs of three runs each (one plain, one bold, one italic), and tables of
  five columns.
* **Build** creates the document from in-memory data, including the time to prepare that data.
* **Extract text** reads the text of all paragraphs.
* **Total** is build + save + open + extract.

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/_static/benchmarks-dark.svg">
    <img alt="Bar charts of the docxium speed-up over python-docx for build, text extraction and total, with the measured times beside each bar." src="docs/_static/benchmarks-light.svg" width="880">
  </picture>
</p>

Save and open are performed by python-docx in both cases; their ratios lie between 0.93× and
1.05×. Timings for every phase are listed in [docs/benchmarks.rst](docs/benchmarks.rst).

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- LICENSE -->
## License

Distributed under the MIT License. See [`LICENSE`](LICENSE) for more information. docxium
depends on python-docx, which is also MIT licensed.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- MARKDOWN LINKS & IMAGES -->
[license-shield]: https://img.shields.io/github/license/cmostw/docxium.svg?style=for-the-badge
[license-url]: https://github.com/cmostw/docxium/blob/master/LICENSE
[Rust-badge]: https://img.shields.io/badge/Rust-000000?style=for-the-badge&logo=rust&logoColor=white
[Rust-url]: https://www.rust-lang.org/
[Python-badge]: https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white
[Python-url]: https://www.python.org/
