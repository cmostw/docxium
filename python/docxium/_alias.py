"""Forwarding of python-docx's module tree.

``import docxium.shared`` / ``from docxium.enum.text import WD_ALIGN_PARAGRAPH``
resolve to the ``docx.shared`` / ``docx.enum.text`` module objects. Names are
looked up in the installed python-docx at import time, so no list of
re-exports is kept.
"""

from __future__ import annotations

import importlib
import importlib.abc
import importlib.util
import sys

_OWN = {"_core", "_alias", "_batch"}  # submodules that really live in docxium


class _AliasLoader(importlib.abc.Loader):
    def __init__(self, target: str) -> None:
        self._target = target

    def create_module(self, spec):
        return importlib.import_module(self._target)

    def exec_module(self, module) -> None:  # already executed as the docx module
        pass


class _AliasFinder(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if not fullname.startswith("docxium."):
            return None
        rest = fullname[len("docxium.") :]
        if rest.split(".")[0] in _OWN:
            return None
        real = "docx." + rest
        try:
            if importlib.util.find_spec(real) is None:
                return None
        except (ImportError, ValueError):
            return None
        return importlib.util.spec_from_loader(fullname, _AliasLoader(real))


def install() -> None:
    if not any(isinstance(f, _AliasFinder) for f in sys.meta_path):
        # must precede PathFinder: for nested names like docxium.dml.color the
        # parent (an alias of docx.dml) has docx's __path__, so PathFinder would
        # otherwise load a second copy of the module under the docxium name.
        sys.meta_path.insert(0, _AliasFinder())
