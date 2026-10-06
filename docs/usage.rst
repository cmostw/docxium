Usage
=====

Everything python-docx provides is used as the python-docx documentation describes (see the
python-docx section of this site). Only the additions below are specific to docxium. They are
available on the document returned by ``docxium.Document()``, or on an existing
``docx.Document`` after ``docxium.accelerate(doc)``.

.. code-block:: python

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

Batch methods
-------------

.. list-table::
   :header-rows: 1
   :widths: 35 65

   * - Method
     - Behavior
   * - ``add_paragraphs(paragraphs, style=None, alignment=None)``
     - Appends one paragraph per item. ``style`` is a style name or object; ``alignment`` is a
       ``WD_ALIGN_PARAGRAPH`` member or its name; both apply to every item. Equivalent to
       ``add_paragraph`` and ``add_run`` calls.
   * - ``add_headings(texts, level=1, alignment=None)``
     - Same as ``add_paragraphs`` with the ``Title`` style for level 0 and ``Heading N`` for
       levels 1-9, as ``add_heading`` selects.
   * - ``add_table_from_rows(rows, style=None, *, header=False, col_widths=None)``
     - Appends a table and returns it. Cells that are not ``str`` are converted with ``str()``.
       ``header=True`` writes the first row in bold. ``col_widths`` holds one width per column (a
       ``docx.shared.Length`` or an EMU integer); without it the text width is divided evenly, as
       ``add_table`` does.
   * - ``extract_paragraph_texts()``
     - Returns the text of every paragraph in document order, including paragraphs inside
       tables. ``Document.paragraphs`` lists body-level paragraphs only.

Paragraphs and runs
-------------------

A paragraph is a ``str`` or an iterable of runs. A run is a ``str``, ``(text, bold, italic)``, or
``(text, {...})`` where the dict may contain:

.. list-table::
   :header-rows: 1
   :widths: 45 55

   * - Key
     - Value
   * - ``bold``, ``italic``, ``underline``, ``strike``, ``all_caps``, ``small_caps``
     - bool
   * - ``superscript``, ``subscript``
     - bool (not both)
   * - ``size``
     - points, for example ``10.5``
   * - ``color``
     - ``"RRGGBB"`` (a leading ``#`` is accepted)
   * - ``font``
     - font name
   * - ``highlight``
     - color name (``"yellow"``, ``"darkBlue"``, ``"dark_blue"``) or a ``WD_COLOR_INDEX`` member

Invalid input raises ``TypeError`` or ``ValueError``.

Output
------

The written XML equals python-docx's for the equivalent calls, with one difference: an empty table
cell is written as ``<w:p/>``, while python-docx writes ``<w:p><w:r/></w:p>``.
