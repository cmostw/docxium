Installation
============

.. code-block:: sh

   pip install docxium

python-docx is installed as a dependency.

Wheels are provided for CPython 3.9 and newer on Linux (x86_64, aarch64), Windows (x64) and
macOS (x86_64, arm64). On other platforms, ``pip`` builds the package from the source
distribution, which requires the Rust toolchain described below.

Building from source
--------------------

Prerequisites
~~~~~~~~~~~~~

* Python 3.9 or newer
* A stable Rust toolchain, 1.85 or newer (`rustup <https://rustup.rs/>`_)
* `maturin <https://github.com/PyO3/maturin>`_

Build and install
~~~~~~~~~~~~~~~~~

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
