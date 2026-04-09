from __future__ import annotations

import tkinter as tk

import app
from tests.tk_test_utils import TkTestCase, destroy_widget


class AutoexpandTests(TkTestCase):
    def test_refresh_autoexpand_resizes_text_after_programmatic_insert(self) -> None:
        widget = tk.Text(self.root, width=40, height=3, wrap="word")
        widget.pack()
        self.addCleanup(destroy_widget, widget)

        app._attach_autoexpand(widget, 3, 20)
        widget.insert("1.0", "Linea 1\nLinea 2\nLinea 3\nLinea 4\nLinea 5")
        app._refresh_autoexpand(widget)
        self.root.update_idletasks()

        self.assertGreaterEqual(int(widget.cget("height")), 5)


if __name__ == "__main__":
    import unittest

    unittest.main()
