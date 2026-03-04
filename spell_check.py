"""
Spell checker helper for Tkinter Text widgets.

Design goals:
- Never break app startup if dependency is missing.
- Debounced checks while typing.
- Right-click suggestions on misspelled words.
"""

from __future__ import annotations

import re
import tkinter as tk
from typing import Iterable

try:
    from spellchecker import SpellChecker
except Exception:
    SpellChecker = None  # type: ignore[assignment]


_checker = None
_WORD_PATTERN = re.compile(
    r"\b[a-zA-Z\u00e1\u00e9\u00ed\u00f3\u00fa\u00fc\u00f1"
    r"\u00c1\u00c9\u00cd\u00d3\u00da\u00dc\u00d1]{3,}\b"
)


def _checker_available() -> bool:
    return SpellChecker is not None


def _get_checker():
    global _checker
    if not _checker_available():
        return None
    if _checker is None:
        _checker = SpellChecker(language="es")
    return _checker


def _normalize_words(words: Iterable[str]) -> list[str]:
    out: list[str] = []
    for w in words:
        w2 = str(w or "").strip()
        if not w2:
            continue
        out.append(w2.lower())
    return out


class SpellCheckHelper:
    def __init__(self, text_widget: tk.Text) -> None:
        self.text = text_widget
        self._after_id: str | None = None
        self.text.tag_configure("misspelled", underline=True, foreground="red")
        self.text.bind("<KeyRelease>", self._on_key_release, add="+")
        self.text.bind("<Button-3>", self._on_right_click, add="+")

    def _on_key_release(self, event) -> None:
        # Skip pure cursor navigation keys
        if event.keysym in {
            "Left",
            "Right",
            "Up",
            "Down",
            "Home",
            "End",
            "Prior",
            "Next",
        }:
            return
        if self._after_id:
            try:
                self.text.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = self.text.after(450, self._check_spelling)

    def _check_spelling(self) -> None:
        self.text.tag_remove("misspelled", "1.0", "end")
        checker = _get_checker()
        if checker is None:
            return

        content = self.text.get("1.0", "end-1c")
        if not content.strip():
            return

        matches = list(_WORD_PATTERN.finditer(content))
        if not matches:
            return

        candidates = [m.group() for m in matches if not m.group().isupper()]
        unknown = checker.unknown(_normalize_words(candidates))

        for match in matches:
            word = match.group()
            if word.isupper():
                continue
            if word.lower() in unknown:
                self.text.tag_add(
                    "misspelled",
                    f"1.0 + {match.start()} chars",
                    f"1.0 + {match.end()} chars",
                )

    def _on_right_click(self, event) -> None:
        menu = tk.Menu(self.text, tearoff=0)
        checker = _get_checker()
        index = self.text.index(f"@{event.x},{event.y}")

        if checker is not None and "misspelled" in self.text.tag_names(index):
            word_start = self.text.index(f"{index} wordstart")
            word_end = self.text.index(f"{index} wordend")
            word = self.text.get(word_start, word_end).strip()

            if word:
                try:
                    suggestions = sorted(checker.candidates(word.lower()) or [])
                except Exception:
                    suggestions = []
                suggestions = suggestions[:6]

                if suggestions:
                    for suggestion in suggestions:
                        # Keep title case when original starts uppercase.
                        if word[:1].isupper():
                            suggestion = suggestion.capitalize()
                        menu.add_command(
                            label=suggestion,
                            command=lambda s=suggestion, ws=word_start, we=word_end: self._replace(ws, we, s),
                        )
                else:
                    menu.add_command(label="Sin sugerencias", state="disabled")

                menu.add_separator()
                menu.add_command(
                    label="Ignorar",
                    command=lambda ws=word_start, we=word_end: self.text.tag_remove("misspelled", ws, we),
                )
                menu.add_command(
                    label="Agregar al diccionario",
                    command=lambda w=word: self._add_word(w),
                )
                menu.add_separator()

        menu.add_command(label="Cortar", command=lambda: self.text.event_generate("<<Cut>>"))
        menu.add_command(label="Copiar", command=lambda: self.text.event_generate("<<Copy>>"))
        menu.add_command(label="Pegar", command=lambda: self.text.event_generate("<<Paste>>"))

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _replace(self, start: str, end: str, word: str) -> None:
        self.text.delete(start, end)
        self.text.insert(start, word)
        self.text.tag_remove("misspelled", start, f"{start} + {len(word)} chars")

    def _add_word(self, word: str) -> None:
        checker = _get_checker()
        if checker is None:
            return
        try:
            checker.word_frequency.load_words([word.lower()])
        except Exception:
            return
        self._check_spelling()


def attach_spell_checker(text_widget: tk.Text):
    """
    Attach spell checker behavior to a tk.Text widget.
    Returns helper instance or None when unavailable.
    """
    if not _checker_available():
        return None
    try:
        if str(text_widget.cget("state")) == "disabled":
            return None
    except Exception:
        return None

    existing = getattr(text_widget, "_spell_helper", None)
    if existing is not None:
        return existing
    helper = SpellCheckHelper(text_widget)
    setattr(text_widget, "_spell_helper", helper)
    return helper

