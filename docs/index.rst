docxium
=======

python-docx extended with batch operations implemented in Rust.

`python-docx <https://github.com/python-openxml/python-docx>`_ adds content to a document one
call at a time. With 50,000 paragraphs of three runs each, building the document takes 39.5 s and
reading all paragraph text back takes 3.5 s (see :doc:`benchmarks`).

docxium keeps python-docx as the document model and adds batch methods. A batch method receives a
whole list of paragraphs or table rows in one call. Rust generates the WordprocessingML for the
list, and python-docx's own parser inserts it into the document, so the result is an ordinary
python-docx ``Document``.

* ``import docxium as docx`` exposes the batch methods. Names and submodules that docxium does not
  define (``docxium.shared``, ``docxium.enum.text``, ...) are forwarded to python-docx.
* Methods that python-docx already provides are python-docx's and are not changed.
* Opening and saving are performed by python-docx and are not changed.

.. toctree::
   :caption: docxium
   :maxdepth: 1

   installation
   usage
   reference
   benchmarks
   attribution

.. toctree::
   :caption: python-docx
   :maxdepth: 1

   python-docx/index
