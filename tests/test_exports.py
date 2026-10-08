"""The package root must actually export every name it advertises.

A name listed in ``__all__`` but never imported is invisible until someone
tries to use it, because a string in ``__all__`` is not checked by any tool.
This walks the list and fails loudly instead.
"""

from __future__ import annotations

import sheetdelta


def test_every_advertised_name_exists():
    missing = [name for name in sheetdelta.__all__ if not hasattr(sheetdelta, name)]
    assert missing == []


def test_the_new_types_are_exported():
    for name in ("Shift", "Table", "TableRef", "render_summary", "CellIndex"):
        assert hasattr(sheetdelta, name), name


def test_the_0_5_types_are_exported():
    for name in ("RowMove", "VolatileRef", "render_markdown"):
        assert hasattr(sheetdelta, name), name
