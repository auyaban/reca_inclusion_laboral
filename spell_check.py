"""
Corrector ortográfico para widgets tk.Text.
Usa pyspellchecker con diccionario en español.
"""
import re
import tkinter as tk
from spellchecker import SpellChecker

_checker = None  # Singleton lazy — el diccionario se carga solo la primera vez

# Palabras de al menos 3 letras (incluye caracteres acentuados del español)
_WORD_PATTERN = re.compile(r'\b[a-záéíóúüñA-ZÁÉÍÓÚÜÑ]{3,}\b')


def _get_checker() -> SpellChecker:
    global _checker
    if _checker is None:
        _checker = SpellChecker(language="es")
    return _checker


class SpellCheckHelper:
    """Adjunta corrección ortográfica a un widget tk.Text."""

    def __init__(self, text_widget: tk.Text) -> None:
        self.text = text_widget
        self._after_id = None
        # Subrayado rojo para palabras mal escritas
        self.text.tag_configure("misspelled", underline=True, foreground="red")
        self.text.bind("<KeyRelease>", self._on_key_release, add="+")
        self.text.bind("<Button-3>", self._on_right_click, add="+")

    # ------------------------------------------------------------------
    # Verificación ortográfica
    # ------------------------------------------------------------------

    def _on_key_release(self, event) -> None:
        # No redibujar al navegar con el teclado
        if event.keysym in ("Left", "Right", "Up", "Down",
                             "Home", "End", "Prior", "Next"):
            return
        if self._after_id:
            self.text.after_cancel(self._after_id)
        # Debounce: esperar 600 ms sin escribir antes de revisar
        self._after_id = self.text.after(600, self._check_spelling)

    def _check_spelling(self) -> None:
        self.text.tag_remove("misspelled", "1.0", "end")
        content = self.text.get("1.0", "end-1c")
        if not content.strip():
            return
        checker = _get_checker()
        for match in _WORD_PATTERN.finditer(content):
            word = match.group()
            if word.isupper():          # saltar acrónimos (NASA, EPS, etc.)
                continue
            if checker.unknown([word.lower()]):
                self.text.tag_add(
                    "misspelled",
                    f"1.0 + {match.start()} chars",
                    f"1.0 + {match.end()} chars",
                )

    # ------------------------------------------------------------------
    # Menú contextual (clic derecho)
    # ------------------------------------------------------------------

    def _on_right_click(self, event) -> None:
        index = self.text.index(f"@{event.x},{event.y}")
        menu = tk.Menu(self.text, tearoff=0)

        if "misspelled" in self.text.tag_names(index):
            word_start = self.text.index(f"{index} wordstart")
            word_end = self.text.index(f"{index} wordend")
            word = self.text.get(word_start, word_end).strip()

            if word:
                checker = _get_checker()
                suggestions = sorted(checker.candidates(word.lower()) or [])[:5]

                if suggestions:
                    for suggestion in suggestions:
                        menu.add_command(
                            label=suggestion,
                            command=lambda s=suggestion,
                                           ws=word_start,
                                           we=word_end: self._replace(ws, we, s),
                        )
                else:
                    menu.add_command(label="Sin sugerencias", state="disabled")

                menu.add_separator()
                menu.add_command(
                    label="Ignorar",
                    command=lambda ws=word_start, we=word_end:
                        self.text.tag_remove("misspelled", ws, we),
                )
                menu.add_command(
                    label="Agregar al diccionario",
                    command=lambda w=word: self._add_word(w),
                )
                menu.add_separator()

        menu.add_command(label="Cortar",
                         command=lambda: self.text.event_generate("<<Cut>>"))
        menu.add_command(label="Copiar",
                         command=lambda: self.text.event_generate("<<Copy>>"))
        menu.add_command(label="Pegar",
                         command=lambda: self.text.event_generate("<<Paste>>"))

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ------------------------------------------------------------------
    # Utilidades
    # ------------------------------------------------------------------

    def _replace(self, start: str, end: str, word: str) -> None:
        self.text.delete(start, end)
        self.text.insert(start, word)
        self.text.tag_remove("misspelled", start, f"{start} + {len(word)} chars")

    def _add_word(self, word: str) -> None:
        _get_checker().word_frequency.load_words([word.lower()])
        self._check_spelling()


def attach_spell_checker(text_widget: tk.Text):
    """
    Adjunta el corrector ortográfico en español a un widget tk.Text.
    Los widgets con state="disabled" se ignoran (son de solo lectura).
    Retorna el SpellCheckHelper o None.
    """
    if str(text_widget.cget("state")) == "disabled":
        return None
    return SpellCheckHelper(text_widget)
