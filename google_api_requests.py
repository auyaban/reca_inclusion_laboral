"""
google_api_requests.py — Reintentos con backoff para Google API.

Responsabilidades:
  - execute_google_request_with_retry: reintentos para requests idempotentes
  - execute_google_create_with_confirmation: reintentos para creates no-idempotentes
    (verifica si el recurso ya fue creado antes de reintentar)
  - Clasificación de errores retryable vs. no-retryable (4xx vs 5xx/timeout)
  - Backoff exponencial con jitter + respeto del header Retry-After

Depende de: nada interno
Usado por: google_sheets_client.py, drive_upload.py

IMPORTANTE: No usar execute_google_request_with_retry para uploads resumibles
ni para creates que no sean idempotentes — usar execute_google_create_with_confirmation.
"""
import random
import socket
import time
from urllib.error import URLError


RETRYABLE_GOOGLE_STATUS_CODES = {429, 500, 502, 503, 504}
NON_RETRYABLE_GOOGLE_STATUS_CODES = {400, 401, 403, 404}


def get_google_api_status_code(exception):
    response = getattr(exception, "resp", None)
    candidates = [
        getattr(response, "status", None),
        getattr(response, "status_code", None),
    ]
    if isinstance(response, dict):
        candidates.extend([response.get("status"), response.get("status_code")])
    for candidate in candidates:
        try:
            if candidate is not None:
                return int(candidate)
        except (TypeError, ValueError):
            continue
    return None


def get_google_retry_after_seconds(exception):
    response = getattr(exception, "resp", None)
    header_sources = [response, getattr(response, "headers", None)]
    for headers in header_sources:
        if not headers or not hasattr(headers, "get"):
            continue
        raw = headers.get("retry-after") or headers.get("Retry-After")
        if raw in (None, ""):
            continue
        try:
            return max(0.0, float(raw))
        except (TypeError, ValueError):
            continue
    return None


def is_retryable_google_request_error(exception):
    status_code = get_google_api_status_code(exception)
    if status_code is not None:
        if status_code in NON_RETRYABLE_GOOGLE_STATUS_CODES:
            return False
        if status_code in RETRYABLE_GOOGLE_STATUS_CODES or status_code >= 500:
            return True

    if isinstance(exception, (TimeoutError, socket.timeout, ConnectionError, URLError, OSError)):
        return True

    reason = getattr(exception, "reason", None)
    if isinstance(reason, (TimeoutError, socket.timeout, ConnectionError, URLError, OSError)):
        return True
    return False


def _compute_retry_delay(exception, attempt, *, initial_delay, max_delay):
    retry_after = get_google_retry_after_seconds(exception)
    delay = retry_after
    if delay is None:
        delay = min(max_delay, initial_delay * (2 ** (attempt - 1)))
        delay *= 1.0 + (random.random() * 0.25)
    return max(0.0, float(delay))


def _log_retry(logger, message):
    if logger is None:
        return
    try:
        logger(message)
    except Exception:
        pass


def execute_google_request_with_retry(
    request,
    *,
    operation_name="google_api_request",
    max_attempts=5,
    initial_delay=1.0,
    max_delay=8.0,
    logger=None,
):
    """Execute a reusable Google API request with retry.

    Use this helper only for request objects that are safe to execute more than
    once. Do not use it for non-idempotent create/copy requests or resumable
    uploads with internal state.
    """
    if request is None:
        raise ValueError("request is required")

    for attempt in range(1, max_attempts + 1):
        try:
            return request.execute()
        except Exception as exc:
            if attempt >= max_attempts or not is_retryable_google_request_error(exc):
                raise

            delay = _compute_retry_delay(
                exc,
                attempt,
                initial_delay=initial_delay,
                max_delay=max_delay,
            )
            _log_retry(
                logger,
                f"RETRY_GOOGLE_API operation={operation_name} "
                f"attempt={attempt}/{max_attempts} wait={delay:.2f}s error={exc}",
            )
            time.sleep(delay)


def execute_google_create_with_confirmation(
    request_factory,
    confirm_created,
    *,
    operation_name="google_create_request",
    max_attempts=5,
    initial_delay=1.0,
    max_delay=8.0,
    logger=None,
):
    """Execute a non-idempotent create/copy request with confirmation.

    The request is rebuilt on every attempt. After any retryable failure, the
    caller-provided confirmation hook checks whether the resource was actually
    created by the ambiguous attempt before another request is issued.
    """
    if not callable(request_factory):
        raise ValueError("request_factory must be callable")
    if not callable(confirm_created):
        raise ValueError("confirm_created must be callable")

    last_error = None
    for attempt in range(1, max_attempts + 1):
        request = request_factory()
        if request is None:
            raise ValueError("request_factory must return a request")
        try:
            return request.execute()
        except Exception as exc:
            last_error = exc
            if not is_retryable_google_request_error(exc):
                raise

            confirmed = None
            try:
                confirmed = confirm_created()
            except Exception as confirm_exc:
                _log_retry(
                    logger,
                    f"WARN_GOOGLE_CREATE_CONFIRM_FAILED operation={operation_name} "
                    f"attempt={attempt}/{max_attempts} error={confirm_exc}",
                )
            if confirmed:
                _log_retry(
                    logger,
                    f"CONFIRMED_GOOGLE_CREATE operation={operation_name} "
                    f"attempt={attempt}/{max_attempts}",
                )
                return confirmed

            if attempt >= max_attempts:
                break

            delay = _compute_retry_delay(
                exc,
                attempt,
                initial_delay=initial_delay,
                max_delay=max_delay,
            )
            _log_retry(
                logger,
                f"RETRY_GOOGLE_CREATE operation={operation_name} "
                f"attempt={attempt}/{max_attempts} wait={delay:.2f}s error={exc}",
            )
            time.sleep(delay)

    try:
        confirmed = confirm_created()
    except Exception as confirm_exc:
        _log_retry(
            logger,
            f"WARN_GOOGLE_CREATE_CONFIRM_FAILED operation={operation_name} "
            f"attempt=final error={confirm_exc}",
        )
        confirmed = None
    if confirmed:
        _log_retry(logger, f"CONFIRMED_GOOGLE_CREATE operation={operation_name} attempt=final")
        return confirmed
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"No se pudo completar la operacion Google: {operation_name}")
