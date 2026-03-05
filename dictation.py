"""
Dictation helper for Tkinter Text widgets using OpenAI STT behind Supabase Edge Functions.

Security model:
- OpenAI key is never present in the desktop app.
- App sends audio to a protected Supabase Edge Function with user JWT.
"""

from __future__ import annotations

import json
import os
import queue
import threading
import time
import uuid
import wave
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Optional

import tkinter as tk

try:
    import sounddevice as sd
except Exception:
    sd = None  # type: ignore[assignment]

from formularios.common import _load_env_file, _load_supabase_credentials


DICTATION_DIR_NAME = "dictation_audio"
DEFAULT_SAMPLE_RATE = 16000
DEFAULT_CHANNELS = 1
DEFAULT_LANGUAGE = "es"
DEFAULT_FUNCTION_NAME = "dictate-transcribe"
DEFAULT_TIMEOUT_SECONDS = 360
MAX_AUDIO_MB = 25

_ACTIVE_RECORDING_LOCK = threading.Lock()
_ACTIVE_RECORDING_ID = ""


def _now():
    return time.time()


def _new_id() -> str:
    return uuid.uuid4().hex


def _get_dictation_dir() -> str:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        base = os.path.join(local_app_data, "RECA", DICTATION_DIR_NAME)
    else:
        base = os.path.join(os.getcwd(), ".cache", DICTATION_DIR_NAME)
    os.makedirs(base, exist_ok=True)
    return base


def _make_audio_path(form_id: str, field_id: str) -> str:
    safe_form = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in str(form_id or "form"))
    safe_field = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in str(field_id or "field"))
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    file_name = f"{safe_form}_{safe_field}_{stamp}_{_new_id()}.wav"
    return os.path.join(_get_dictation_dir(), file_name)


def _safe_remove(path: str) -> None:
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def _extract_error_message(exc: Exception) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        try:
            body = exc.read().decode("utf-8", errors="replace")
            payload = json.loads(body) if body else {}
            if isinstance(payload, dict):
                err = payload.get("error")
                if isinstance(err, dict):
                    msg = str(err.get("message") or "").strip()
                    if msg:
                        return msg
                for key in ("message", "error_description", "error", "msg"):
                    value = payload.get(key)
                    if value:
                        return str(value)
            if body:
                return body
        except Exception:
            pass
    return str(exc)


def _read_config():
    env = _load_env_file(".env")
    fn_name = str(env.get("DICTATION_FUNCTION_NAME") or DEFAULT_FUNCTION_NAME).strip() or DEFAULT_FUNCTION_NAME
    language = str(env.get("DICTATION_LANGUAGE") or DEFAULT_LANGUAGE).strip() or DEFAULT_LANGUAGE
    return {
        "function_name": fn_name,
        "language": language,
    }


def _encode_multipart(fields: dict, file_field: str, file_name: str, file_bytes: bytes, content_type: str):
    boundary = f"----RECADictationBoundary{uuid.uuid4().hex}"
    lines = []
    for key, value in fields.items():
        lines.append(f"--{boundary}\r\n".encode("utf-8"))
        lines.append(
            f'Content-Disposition: form-data; name="{key}"\r\n\r\n{value}\r\n'.encode("utf-8")
        )
    lines.append(f"--{boundary}\r\n".encode("utf-8"))
    lines.append(
        (
            f'Content-Disposition: form-data; name="{file_field}"; filename="{file_name}"\r\n'
            f"Content-Type: {content_type}\r\n\r\n"
        ).encode("utf-8")
    )
    lines.append(file_bytes)
    lines.append(b"\r\n")
    lines.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(lines), boundary


class _Recorder:
    def __init__(self, file_path: str, sample_rate: int = DEFAULT_SAMPLE_RATE, channels: int = DEFAULT_CHANNELS):
        if sd is None:
            raise RuntimeError("Dependencia de audio no disponible (sounddevice).")
        self.file_path = file_path
        self.sample_rate = int(sample_rate)
        self.channels = int(channels)
        self._queue: queue.Queue[bytes] = queue.Queue()
        self._stop_event = threading.Event()
        self._writer_thread: Optional[threading.Thread] = None
        self._stream = None
        self._started = False

    def _callback(self, indata, frames, _time_info, status):
        _ = frames
        if status:
            # Do not fail hard on callback status; keep best-effort recording.
            pass
        try:
            self._queue.put_nowait(indata.copy().tobytes())
        except Exception:
            pass

    def _writer_loop(self):
        with wave.open(self.file_path, "wb") as wav_file:
            wav_file.setnchannels(self.channels)
            wav_file.setsampwidth(2)  # int16
            wav_file.setframerate(self.sample_rate)
            while not self._stop_event.is_set() or not self._queue.empty():
                try:
                    chunk = self._queue.get(timeout=0.25)
                except queue.Empty:
                    continue
                if chunk:
                    wav_file.writeframes(chunk)

    def start(self):
        if self._started:
            return
        self._writer_thread = threading.Thread(target=self._writer_loop, daemon=True)
        self._writer_thread.start()
        self._stream = sd.InputStream(
            samplerate=self.sample_rate,
            channels=self.channels,
            dtype="int16",
            callback=self._callback,
        )
        self._stream.start()
        self._started = True

    def stop(self):
        if not self._started:
            return
        try:
            if self._stream is not None:
                self._stream.stop()
                self._stream.close()
        finally:
            self._stream = None
            self._stop_event.set()
            if self._writer_thread is not None:
                self._writer_thread.join(timeout=5.0)
            self._started = False


@dataclass
class RecordingHandle:
    id: str
    file_path: str
    started_at: float
    form_id: str
    field_id: str
    _recorder: _Recorder
    _stopped: bool = False


@dataclass
class DictationResult:
    ok: bool
    text: str = ""
    elapsed_ms: int = 0
    error_code: str = ""
    error_message: str = ""


def _claim_global_recording(handle_id: str) -> bool:
    global _ACTIVE_RECORDING_ID
    with _ACTIVE_RECORDING_LOCK:
        if _ACTIVE_RECORDING_ID and _ACTIVE_RECORDING_ID != handle_id:
            return False
        _ACTIVE_RECORDING_ID = handle_id
        return True


def _release_global_recording(handle_id: str) -> None:
    global _ACTIVE_RECORDING_ID
    with _ACTIVE_RECORDING_LOCK:
        if _ACTIVE_RECORDING_ID == handle_id:
            _ACTIVE_RECORDING_ID = ""


def start_recording(field_id: str, form_id: str) -> RecordingHandle:
    if sd is None:
        raise RuntimeError("Dictado no disponible: instala dependencia de audio (sounddevice).")
    path = _make_audio_path(form_id=form_id, field_id=field_id)
    recorder = _Recorder(path)
    recorder.start()
    return RecordingHandle(
        id=_new_id(),
        file_path=path,
        started_at=_now(),
        form_id=str(form_id or ""),
        field_id=str(field_id or ""),
        _recorder=recorder,
    )


def _stop_handle_if_needed(handle: RecordingHandle):
    if handle and not handle._stopped:
        handle._recorder.stop()
        handle._stopped = True


def _post_audio_to_edge(handle: RecordingHandle, jwt_token: str) -> DictationResult:
    started = _now()
    if not os.path.exists(handle.file_path):
        return DictationResult(ok=False, error_code="audio_missing", error_message="No se encontro el audio.")

    size_bytes = os.path.getsize(handle.file_path)
    if size_bytes <= 0:
        return DictationResult(ok=False, error_code="audio_empty", error_message="El audio grabado esta vacio.")
    if size_bytes > (MAX_AUDIO_MB * 1024 * 1024):
        return DictationResult(
            ok=False,
            error_code="audio_too_large",
            error_message=f"El audio supera el limite de {MAX_AUDIO_MB}MB.",
        )

    try:
        supabase_url, supabase_key = _load_supabase_credentials(".env")
    except Exception as exc:
        return DictationResult(
            ok=False,
            error_code="supabase_config",
            error_message=f"No se pudo cargar configuracion Supabase: {exc}",
        )

    cfg = _read_config()
    fn_name = cfg.get("function_name") or DEFAULT_FUNCTION_NAME
    language = cfg.get("language") or DEFAULT_LANGUAGE
    url = f"{supabase_url.rstrip('/')}/functions/v1/{fn_name}"

    with open(handle.file_path, "rb") as in_file:
        audio_bytes = in_file.read()

    fields = {
        "form_id": handle.form_id,
        "field_id": handle.field_id,
        "language": language,
    }
    body, boundary = _encode_multipart(
        fields=fields,
        file_field="audio_file",
        file_name=os.path.basename(handle.file_path),
        file_bytes=audio_bytes,
        content_type="audio/wav",
    )

    headers = {
        "apikey": supabase_key,
        "Authorization": f"Bearer {str(jwt_token or '').strip()}",
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    }
    request = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=DEFAULT_TIMEOUT_SECONDS) as response:
            raw = response.read().decode("utf-8", errors="replace")
        payload = json.loads(raw) if raw else {}
    except Exception as exc:
        return DictationResult(
            ok=False,
            elapsed_ms=int((_now() - started) * 1000),
            error_code="edge_request_failed",
            error_message=_extract_error_message(exc),
        )

    ok = bool(payload.get("ok", False))
    text = str(payload.get("text") or "")
    if not ok or not text.strip():
        err = payload.get("error")
        message = ""
        if isinstance(err, dict):
            message = str(err.get("message") or "").strip()
        if not message:
            message = str(payload.get("message") or "No se recibio transcripcion.")
        return DictationResult(
            ok=False,
            elapsed_ms=int((_now() - started) * 1000),
            error_code=str((err or {}).get("code") if isinstance(err, dict) else "transcription_failed"),
            error_message=message,
        )

    return DictationResult(ok=True, text=text, elapsed_ms=int((_now() - started) * 1000))


def stop_and_transcribe(handle: RecordingHandle, jwt_token: str) -> DictationResult:
    try:
        _stop_handle_if_needed(handle)
    except Exception as exc:
        _release_global_recording(handle.id)
        return DictationResult(ok=False, error_code="record_stop_failed", error_message=str(exc))

    token = str(jwt_token or "").strip()
    if not token:
        _release_global_recording(handle.id)
        return DictationResult(
            ok=False,
            error_code="missing_session",
            error_message="No hay sesion valida para usar dictado.",
        )

    result = _post_audio_to_edge(handle, token)
    if result.ok:
        _safe_remove(handle.file_path)
    _release_global_recording(handle.id)
    return result


def cancel_recording(handle: RecordingHandle) -> None:
    try:
        _stop_handle_if_needed(handle)
    except Exception:
        pass
    _safe_remove(handle.file_path)
    _release_global_recording(handle.id)


def cleanup_stale_audio(ttl_hours: int = 24) -> int:
    deleted = 0
    cutoff = _now() - max(1, int(ttl_hours)) * 3600
    folder = _get_dictation_dir()
    try:
        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            if not os.path.isfile(path):
                continue
            if not name.lower().endswith(".wav"):
                continue
            try:
                mtime = os.path.getmtime(path)
                if mtime < cutoff:
                    os.remove(path)
                    deleted += 1
            except Exception:
                continue
    except Exception:
        return deleted
    return deleted


class DictationTextHelper:
    def __init__(
        self,
        text_widget: tk.Text,
        *,
        form_id: str,
        field_id: str,
        session_provider: Callable[[], str],
        log_fn: Optional[Callable[[str], None]] = None,
    ):
        self.text = text_widget
        self.form_id = str(form_id or "")
        self.field_id = str(field_id or "")
        self.session_provider = session_provider
        self.log_fn = log_fn
        self._handle: Optional[RecordingHandle] = None
        self._tick_after_id = None
        self._worker = None
        self._is_recording = False
        self._controls_parent = self.text.master
        self._controls = tk.Frame(self._controls_parent, bg="#F2F2F2", bd=1, relief="solid")
        self._button = tk.Button(
            self._controls,
            text="Dictar",
            width=8,
            command=self._on_toggle,
            cursor="hand2",
        )
        self._button.pack(side="left", padx=(4, 2), pady=2)
        self._status = tk.Label(
            self._controls,
            text="Listo",
            bg="#F2F2F2",
            fg="#333333",
            font=("Arial", 8),
        )
        self._status.pack(side="left", padx=(0, 6), pady=2)
        self._controls.place(x=0, y=0)
        self.text.bind("<Configure>", self._place_controls, add="+")
        self._controls_parent.bind("<Configure>", self._place_controls, add="+")
        self.text.bind("<Destroy>", self._on_destroy, add="+")
        self.text.after(0, self._place_controls)
        self._sync_enabled_state()

    def _log(self, message: str):
        if callable(self.log_fn):
            try:
                self.log_fn(message)
            except Exception:
                pass

    def _set_status(self, text: str, fg: str = "#333333"):
        self._status.config(text=text, fg=fg)

    def _sync_enabled_state(self):
        try:
            state = str(self.text.cget("state"))
        except Exception:
            state = "normal"
        enabled = (state != "disabled") and (sd is not None)
        self._button.config(state="normal" if enabled else "disabled")
        if not enabled and sd is None:
            self._set_status("Audio no disponible", "#A40000")

    def _on_destroy(self, _event):
        if self._handle is not None:
            cancel_recording(self._handle)
            self._handle = None
        if self._tick_after_id:
            try:
                self.text.after_cancel(self._tick_after_id)
            except Exception:
                pass
            self._tick_after_id = None
        try:
            self._controls.destroy()
        except Exception:
            pass

    def _place_controls(self, _event=None):
        try:
            if not self.text.winfo_exists() or not self._controls.winfo_exists():
                return
            parent = self._controls_parent
            parent_w = max(1, parent.winfo_width())
            parent_h = max(1, parent.winfo_height())
            text_x = self.text.winfo_x()
            text_y = self.text.winfo_y()
            text_w = self.text.winfo_width()
            text_h = self.text.winfo_height()
            controls_w = max(1, self._controls.winfo_reqwidth())
            controls_h = max(1, self._controls.winfo_reqheight())

            x = text_x + text_w + 6
            y = text_y

            if x + controls_w > parent_w:
                x_left = text_x - controls_w - 6
                if x_left >= 0:
                    x = x_left
                else:
                    x = min(max(0, text_x + text_w - controls_w), max(0, parent_w - controls_w))
                    y = min(text_y + text_h + 2, max(0, parent_h - controls_h))

            y = max(0, min(y, max(0, parent_h - controls_h)))
            self._controls.place(x=int(x), y=int(y))
        except Exception:
            return

    def _on_toggle(self):
        self._sync_enabled_state()
        if self._is_recording:
            self._stop_and_transcribe_async()
        else:
            self._start_recording()

    def _start_recording(self):
        try:
            handle = start_recording(field_id=self.field_id, form_id=self.form_id)
        except Exception as exc:
            self._set_status("Error", "#A40000")
            self._log(f"[DICTATION] No se pudo iniciar grabacion: {exc}")
            return
        if not _claim_global_recording(handle.id):
            cancel_recording(handle)
            self._set_status("Otro dictado activo", "#A40000")
            return
        self._handle = handle
        self._is_recording = True
        self._button.config(text="Detener")
        self._set_status("Grabando 00:00", "#B35300")
        self._log(f"[DICTATION] Grabando form={self.form_id} field={self.field_id} file={handle.file_path}")
        self._tick_timer()

    def _tick_timer(self):
        if not self._is_recording or not self._handle:
            return
        elapsed = int(max(0, _now() - self._handle.started_at))
        mm = elapsed // 60
        ss = elapsed % 60
        label = f"Grabando {mm:02d}:{ss:02d}"
        if elapsed >= 600:
            label += " (largo)"
        self._set_status(label, "#B35300")
        self._tick_after_id = self.text.after(1000, self._tick_timer)

    def _stop_and_transcribe_async(self):
        if not self._handle:
            self._reset_controls()
            return
        self._is_recording = False
        if self._tick_after_id:
            try:
                self.text.after_cancel(self._tick_after_id)
            except Exception:
                pass
            self._tick_after_id = None
        self._button.config(state="disabled", text="Procesando")
        self._set_status("Procesando...", "#005A9C")
        handle = self._handle
        self._handle = None

        def _work():
            token = ""
            try:
                token = str(self.session_provider() or "").strip()
            except Exception:
                token = ""
            result = stop_and_transcribe(handle, token)
            self.text.after(0, lambda: self._on_transcription_result(result))

        self._worker = threading.Thread(target=_work, daemon=True)
        self._worker.start()

    def _on_transcription_result(self, result: DictationResult):
        if result.ok:
            try:
                self.text.insert("insert", result.text)
                self.text.see("insert")
            except Exception:
                pass
            self._set_status("Listo", "#007A2F")
            self._log(
                f"[DICTATION] OK form={self.form_id} field={self.field_id} elapsed_ms={result.elapsed_ms}"
            )
        else:
            self._set_status("Error", "#A40000")
            self._log(
                f"[DICTATION] ERROR form={self.form_id} field={self.field_id} "
                f"code={result.error_code} msg={result.error_message}"
            )
            try:
                from tkinter import messagebox

                messagebox.showerror(
                    "Dictado",
                    f"No se pudo transcribir.\n{result.error_message or 'Error desconocido.'}",
                )
            except Exception:
                pass
        self._reset_controls()

    def _reset_controls(self):
        self._is_recording = False
        self._button.config(text="Dictar", state="normal")
        self._sync_enabled_state()
        if self._status.cget("text") == "Procesando...":
            self._set_status("Listo", "#333333")


def attach_dictation(
    text_widget: tk.Text,
    *,
    form_id: str,
    field_id: str,
    session_provider: Callable[[], str],
    log_fn: Optional[Callable[[str], None]] = None,
):
    try:
        if str(text_widget.cget("state")) == "disabled":
            return None
    except Exception:
        return None

    existing = getattr(text_widget, "_dictation_helper", None)
    if existing is not None:
        return existing
    helper = DictationTextHelper(
        text_widget,
        form_id=form_id,
        field_id=field_id,
        session_provider=session_provider,
        log_fn=log_fn,
    )
    setattr(text_widget, "_dictation_helper", helper)
    return helper
