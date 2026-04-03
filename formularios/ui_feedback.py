from __future__ import annotations

import tkinter as tk


STATE_COLORS = {
    "info": "#1F2A44",
    "loading": "#07B499",
    "success": "#0A7D2E",
    "warning": "#B35300",
    "error": "#B00020",
    "hint": "#6B6B6B",
}

FIELD_ERROR_BG = "#FDE2E2"
PLACEHOLDER_FG = "#6B6B6B"
TEXT_FG = "#222222"


def state_color(state: str, default: str | None = None) -> str:
    key = str(state or "").strip().lower()
    return STATE_COLORS.get(key, default or STATE_COLORS["info"])


def set_semantic_label(label, text: str, *, state: str = "info") -> None:
    if label is None:
        return
    try:
        label.config(text=str(text or ""), fg=state_color(state))
    except Exception:
        return


def ensure_feedback_state(host) -> dict:
    state = getattr(host, "_ui_feedback_state", None)
    if isinstance(state, dict):
        return state
    state = {
        "fields": {},
        "errors": {},
        "group_hints": {},
        "banner_label": None,
    }
    host._ui_feedback_state = state
    return state


def register_banner_label(host, label) -> None:
    ensure_feedback_state(host)["banner_label"] = label


def set_banner(host, text: str, *, state: str = "error", label=None) -> None:
    feedback = ensure_feedback_state(host)
    target = label or feedback.get("banner_label")
    if target is None:
        return
    set_semantic_label(target, text, state=state)


def clear_banner(host, *, label=None) -> None:
    set_banner(host, "", state="info", label=label)


def _field_bucket(host, field_id: str) -> dict:
    feedback = ensure_feedback_state(host)
    bucket = feedback["fields"].setdefault(
        str(field_id or "").strip(),
        {
            "widgets": [],
            "error_labels": [],
        },
    )
    return bucket


def register_field(host, field_id: str, widget, *, error_label=None) -> None:
    bucket = _field_bucket(host, field_id)
    if widget is not None and widget not in bucket["widgets"]:
        bucket["widgets"].append(widget)
        _bind_auto_clear(host, field_id, widget)
    if error_label is not None and error_label not in bucket["error_labels"]:
        bucket["error_labels"].append(error_label)


def register_error_label(host, field_id: str, label) -> None:
    register_field(host, field_id, None, error_label=label)


def _bind_auto_clear(host, field_id: str, widget) -> None:
    if getattr(widget, "_ui_feedback_clear_bound", False):
        return

    def _clear(_event=None):
        clear_field_error(host, field_id)

    for sequence in ("<KeyRelease>", "<FocusIn>", "<<ComboboxSelected>>", "<<Modified>>"):
        try:
            widget.bind(sequence, _clear, add="+")
        except Exception:
            continue
    widget._ui_feedback_clear_bound = True


def _apply_invalid_visual(widget, invalid: bool) -> None:
    if widget is None:
        return
    try:
        klass = str(widget.winfo_class() or "")
    except Exception:
        klass = ""
    if klass in {"TEntry", "TCombobox"}:
        original_style = getattr(widget, "_ui_feedback_original_style", None)
        if original_style is None:
            try:
                original_style = str(widget.cget("style") or "")
            except Exception:
                original_style = ""
            widget._ui_feedback_original_style = original_style
        target_style = f"Invalid.{klass}" if invalid else original_style
        try:
            widget.configure(style=target_style)
        except Exception:
            pass
        return
    if klass in {"Entry", "Text", "DateEntry"}:
        original_bg = getattr(widget, "_ui_feedback_original_bg", None)
        if original_bg is None:
            try:
                original_bg = widget.cget("bg")
            except Exception:
                original_bg = "white"
            widget._ui_feedback_original_bg = original_bg
        try:
            widget.configure(bg=FIELD_ERROR_BG if invalid else original_bg)
        except Exception:
            pass


def set_field_error(host, field_id: str, message: str) -> None:
    bucket = _field_bucket(host, field_id)
    feedback = ensure_feedback_state(host)
    feedback["errors"][str(field_id or "").strip()] = str(message or "")
    for widget in list(bucket.get("widgets") or []):
        _apply_invalid_visual(widget, True)
    labels = list(bucket.get("error_labels") or [])
    for label in labels:
        set_semantic_label(label, message, state="error")


def clear_field_error(host, field_id: str) -> None:
    feedback = ensure_feedback_state(host)
    bucket = feedback.get("fields", {}).get(str(field_id or "").strip()) or {}
    feedback.get("errors", {}).pop(str(field_id or "").strip(), None)
    for widget in list(bucket.get("widgets") or []):
        _apply_invalid_visual(widget, False)
    for label in list(bucket.get("error_labels") or []):
        set_semantic_label(label, "", state="hint")


def clear_field_errors(host, field_ids=None) -> None:
    feedback = ensure_feedback_state(host)
    keys = field_ids
    if keys is None:
        keys = list((feedback.get("fields") or {}).keys())
    for field_id in list(keys or []):
        clear_field_error(host, field_id)


def focus_first_invalid_field(host, field_ids=None) -> bool:
    feedback = ensure_feedback_state(host)
    ordered = field_ids
    if ordered is None:
        ordered = list((feedback.get("errors") or {}).keys())
    for field_id in list(ordered or []):
        bucket = (feedback.get("fields") or {}).get(str(field_id or "").strip()) or {}
        for widget in list(bucket.get("widgets") or []):
            try:
                widget.focus_set()
                return True
            except Exception:
                continue
    return False


def register_group_hint(host, hint_id: str, label) -> None:
    ensure_feedback_state(host)["group_hints"][str(hint_id or "").strip()] = label


def set_group_hint(host, hint_id: str, text: str, *, state: str = "hint") -> None:
    label = ensure_feedback_state(host).get("group_hints", {}).get(str(hint_id or "").strip())
    set_semantic_label(label, text, state=state)


def bind_placeholder(widget, text: str) -> None:
    placeholder = str(text or "").strip()
    if not placeholder or getattr(widget, "_ui_placeholder_bound", False):
        return
    widget._ui_placeholder_text = placeholder
    widget._ui_placeholder_bound = True

    def _set_foreground(color: str) -> None:
        try:
            widget.configure(foreground=color)
        except Exception:
            try:
                widget.configure(fg=color)
            except Exception:
                pass

    def _show_placeholder() -> None:
        current = _get_widget_text(widget)
        if current:
            return
        _set_widget_text(widget, placeholder)
        widget._ui_placeholder_active = True
        _set_foreground(PLACEHOLDER_FG)

    def _clear_placeholder(_event=None) -> None:
        if not getattr(widget, "_ui_placeholder_active", False):
            return
        _set_widget_text(widget, "")
        widget._ui_placeholder_active = False
        _set_foreground(TEXT_FG)

    def _restore_placeholder(_event=None) -> None:
        current = _get_widget_text(widget)
        if current:
            widget._ui_placeholder_active = False
            _set_foreground(TEXT_FG)
            return
        _show_placeholder()

    try:
        widget.bind("<FocusIn>", _clear_placeholder, add="+")
        widget.bind("<FocusOut>", _restore_placeholder, add="+")
        widget.bind("<<ComboboxSelected>>", _restore_placeholder, add="+")
    except Exception:
        return
    _show_placeholder()


def is_placeholder_active(widget) -> bool:
    return bool(getattr(widget, "_ui_placeholder_active", False))


def get_widget_value(widget) -> str:
    if widget is None:
        return ""
    if is_placeholder_active(widget):
        return ""
    return _get_widget_text(widget)


def _get_widget_text(widget) -> str:
    if widget is None:
        return ""
    try:
        return str(widget.get()).strip()
    except Exception:
        return ""


def _set_widget_text(widget, value: str) -> None:
    try:
        widget.delete(0, tk.END)
        widget.insert(0, value)
    except Exception:
        return
