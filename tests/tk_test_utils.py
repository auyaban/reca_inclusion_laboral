from __future__ import annotations

import tkinter as tk
import unittest


def cancel_pending_after_callbacks(root) -> None:
    if root is None:
        return
    try:
        pending = root.tk.call("after", "info")
    except Exception:
        return
    if isinstance(pending, str):
        pending = (pending,) if pending else ()
    for after_id in tuple(pending or ()):
        try:
            root.after_cancel(after_id)
        except Exception:
            continue


def destroy_widget(widget) -> None:
    if widget is None:
        return
    try:
        root = widget.winfo_toplevel()
    except Exception:
        root = None
    cancel_pending_after_callbacks(root)
    try:
        widget.update_idletasks()
    except Exception:
        pass
    try:
        if widget.winfo_exists():
            widget.destroy()
    except Exception:
        pass
    cancel_pending_after_callbacks(root)


class TkTestCase(unittest.TestCase):
    def setUp(self) -> None:
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk no disponible: {exc}")
        self.root.withdraw()
        self.addCleanup(self._cleanup_root)

    def _cleanup_root(self) -> None:
        try:
            cancel_pending_after_callbacks(self.root)
            for child in tuple(self.root.winfo_children()):
                destroy_widget(child)
            cancel_pending_after_callbacks(self.root)
            self.root.update_idletasks()
            self.root.destroy()
        except Exception:
            pass
