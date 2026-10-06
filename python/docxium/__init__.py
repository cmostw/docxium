"""docxium: python-docx extended with batch operations implemented in Rust.

    import docxium as docx

Attributes and submodules that docxium does not define are forwarded to
python-docx, so every python-docx name remains available.
"""

from __future__ import annotations

import docx as _docx

from . import _alias
from ._batch import Document, DocxiumDocument, accelerate

_alias.install()

__version__ = "0.1.0"
docx_version = _docx.__version__

__all__ = ["Document", "DocxiumDocument", "accelerate"]


def __getattr__(name: str):
    try:
        return getattr(_docx, name)
    except AttributeError:
        raise AttributeError(f"module 'docxium' has no attribute {name!r}") from None


def __dir__():
    return sorted(set(globals()) | set(dir(_docx)))
