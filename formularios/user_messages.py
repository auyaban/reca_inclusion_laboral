from __future__ import annotations

# Todo contexto nuevo usado por _log_user_error(...) debe declararse aqui primero.
SUPPORTED_USER_MESSAGE_CONTEXTS = (
    "login",
    "company_search",
    "linked_user_search",
    "autofill_user",
    "database_refresh",
    "sync",
    "case_open",
    "followup_case",
    "section_confirm",
    "finalization",
    "save_sheet",
    "ui_error",
)


def map_exception_to_user_message(context: str, exc) -> str:
    raw = str(exc or "").strip()
    text = raw.lower()
    ctx = str(context or "").strip().lower()
    if ctx not in SUPPORTED_USER_MESSAGE_CONTEXTS:
        ctx = ""

    if not raw:
        return _default_message(ctx)

    if "timeout" in text or "timed out" in text:
        return "La operación tardó demasiado. Inténtalo de nuevo."
    if any(
        token in text
        for token in ("connection", "network", "dns", "name or service not known", "failed to resolve")
    ):
        return "No fue posible conectarse en este momento. Revisa tu conexión e inténtalo de nuevo."
    if "more than one" in text or "más de una empresa" in text:
        return "Se encontraron varios resultados. Usa el NIT para identificar la empresa correcta."
    if "no se encontró" in text or "not found" in text:
        return raw
    if "invalid login credentials" in text or "invalid_grant" in text:
        return "Usuario o contraseña incorrectos."
    if "usuario y contraseña incorrectos" in text or "usuario o contraseña incorrectos" in text:
        return "Usuario o contraseña incorrectos."
    if ctx == "login" and (
        "email not confirmed" in text
        or "email_not_confirmed" in text
        or "confirmar tu correo" in text
    ):
        return "Debes confirmar tu correo antes de iniciar sesión."
    if "drive" in text and ctx in {"sync", "case_open", "followup_case", "database_refresh", "save_sheet"}:
        return "No fue posible comunicarse con Drive en este momento. Inténtalo de nuevo."
    if "supabase" in text and ctx in {
        "company_search",
        "linked_user_search",
        "autofill_user",
        "database_refresh",
        "followup_case",
        "sync",
    }:
        return "No fue posible consultar la base de datos en este momento. Inténtalo de nuevo."
    if "permisos actuales" in text and ctx == "login":
        return "No fue posible validar tu perfil con los permisos actuales."
    if "permission denied" in text and ctx == "linked_user_search":
        return "No fue posible consultar el vinculado con los permisos actuales."

    return _default_message(ctx)


def _default_message(context: str) -> str:
    defaults = {
        "company_search": "No fue posible consultar la empresa. Inténtalo de nuevo.",
        "linked_user_search": "No fue posible buscar al vinculado. Inténtalo de nuevo.",
        "autofill_user": "No fue posible completar los datos automáticamente. Puedes continuar con el diligenciamiento manual.",
        "database_refresh": "No fue posible actualizar la base de datos. Inténtalo de nuevo.",
        "sync": "No fue posible consultar el estado de sincronización.",
        "case_open": "No fue posible abrir el caso en este momento.",
        "followup_case": "No fue posible preparar el caso de seguimiento.",
        "section_confirm": "No fue posible guardar esta sección. Revisa los datos e inténtalo de nuevo.",
        "finalization": "No fue posible finalizar el formulario.",
        "login": "No fue posible iniciar sesión en este momento.",
        "save_sheet": "No fue posible guardar la hoja actual. Inténtalo de nuevo.",
        "ui_error": "No fue posible completar esta acción. Inténtalo de nuevo.",
    }
    return defaults.get(str(context or "").strip().lower(), "No fue posible completar la operación.")
