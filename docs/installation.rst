Installation
============

docxium is built from source; there is no PyPI release.

Prerequisites
-------------

* Python 3.9 or newer
* A stable Rust toolchain, 1.85 or newer (`rustup <https://rustup.rs/>`_)
* `maturin <https://github.com/PyO3/maturin>`_

python-docx is installed as a dependency.

Build and install
-----------------

.. code-block:: sh

   git clone https://github.com/cmostw/docxium.git
   cd docxium
   python -m venv .venv
   pip install maturin
   maturin develop --release

``maturin develop --release`` builds the extension and installs it into the active environment.
To run the tests:

.. code-block:: sh

   python -m unittest discover tests

``maturin build --release`` produces a wheel built against the stable ABI (abi3, Python 3.9+).

Building this documentation
---------------------------

.. code-block:: sh

   pip install -r docs/requirements.txt
   sphinx-build -b html docs docs/_build/html

The API pages import ``docxium``, so the package must be installed in the same environment.
