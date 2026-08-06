"""Supabase-Zugriff fuer NursePlan Pro.

Der Supabase-Client wird absichtlich pro Streamlit-Sitzung gespeichert. Dadurch
teilen sich gleichzeitig angemeldete Nutzer weder Login noch Zugriffstoken.
Die Row-Level-Security-Regeln in Supabase bleiben die letzte Sicherheitslinie.
"""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any

import streamlit as st
from supabase import Client, create_client


class SupabaseConfigurationError(RuntimeError):
    """Die lokale Supabase-Konfiguration fehlt oder ist unsicher."""


class NursePlanDatabaseError(RuntimeError):
    """Eine erwartete Datenbankoperation wurde nicht ausgefuehrt."""


def _secret_value(*names: str) -> str:
    try:
        for name in names:
            value = st.secrets.get(name)
            if value:
                return str(value).strip()
    except Exception as exc:
        raise SupabaseConfigurationError(
            "Die Datei .streamlit/secrets.toml fehlt oder ist ungueltig."
        ) from exc
    return ""


def get_supabase_client() -> Client:
    """Liefert genau einen authentifizierten Client je Streamlit-Sitzung."""
    if "supabase_client" in st.session_state:
        return st.session_state.supabase_client

    url = _secret_value("SUPABASE_URL")
    key = _secret_value(
        "SUPABASE_PUBLISHABLE_KEY",
        "SUPABASE_ANON_KEY",  # Kompatibel mit aelteren Supabase-Projekten.
    )
    if not url or not key:
        raise SupabaseConfigurationError(
            "SUPABASE_URL und SUPABASE_PUBLISHABLE_KEY fehlen."
        )
    if key.startswith("sb_secret_"):
        raise SupabaseConfigurationError(
            "Ein Secret Key darf nicht in der App verwendet werden. "
            "Bitte den Publishable Key verwenden."
        )

    st.session_state.supabase_client = create_client(url, key)
    return st.session_state.supabase_client


def get_supabase_public_config() -> tuple[str, str]:
    """Liefert ausschließlich die für Browser vorgesehenen Supabase-Werte."""
    url = _secret_value("SUPABASE_URL")
    key = _secret_value("SUPABASE_PUBLISHABLE_KEY", "SUPABASE_ANON_KEY")
    if not url or not key or key.startswith("sb_secret_"):
        raise SupabaseConfigurationError("Die öffentliche Supabase-Konfiguration fehlt.")
    return url.rstrip("/"), key


def _remember_user(response: Any) -> bool:
    user = getattr(response, "user", None)
    session = getattr(response, "session", None)
    if user is None or session is None:
        return False

    metadata = getattr(user, "user_metadata", None) or {}
    st.session_state.auth_user_id = str(user.id)
    st.session_state.auth_email = str(getattr(user, "email", "") or "")
    st.session_state.auth_name = str(metadata.get("full_name", "") or "")
    return True


def _refresh_token_from_response(response: Any) -> str:
    session = getattr(response, "session", None)
    return str(getattr(session, "refresh_token", "") or "")


def sign_in_user(email: str, password: str) -> str:
    """Meldet an und gibt den Refresh-Token für eine optionale Browser-Sitzung zurück."""
    response = get_supabase_client().auth.sign_in_with_password(
        {"email": email.strip().lower(), "password": password}
    )
    if not _remember_user(response):
        raise NursePlanDatabaseError("Die Anmeldung konnte nicht abgeschlossen werden.")
    return _refresh_token_from_response(response)


def sign_up_user(full_name: str, email: str, password: str) -> str:
    """Registriert ein Konto und gibt bei sofortiger Anmeldung den Refresh-Token zurück."""
    response = get_supabase_client().auth.sign_up(
        {
            "email": email.strip().lower(),
            "password": password,
            "options": {"data": {"full_name": full_name.strip()}},
        }
    )
    return _refresh_token_from_response(response) if _remember_user(response) else ""


def request_password_reset(email: str, redirect_url: str) -> None:
    """Fordert bei Supabase eine E-Mail zum Zurücksetzen des Passworts an."""
    normalized_email = str(email).strip().lower()
    target_url = str(redirect_url).strip()
    if not normalized_email or "@" not in normalized_email:
        raise NursePlanDatabaseError("Bitte eine gültige E-Mail-Adresse eingeben.")
    if not target_url.startswith(("http://", "https://")):
        raise NursePlanDatabaseError("Die Rückkehr-Adresse der App ist ungültig.")
    get_supabase_client().auth.reset_password_for_email(
        normalized_email,
        {"redirect_to": target_url},
    )


def restore_user_session(refresh_token: str) -> str:
    """Stellt eine Supabase-Sitzung wieder her und liefert den rotierten Token."""
    token = str(refresh_token or "").strip()
    if not token:
        raise NursePlanDatabaseError("Die gespeicherte Sitzung fehlt.")
    response = get_supabase_client().auth.refresh_session(token)
    if not _remember_user(response):
        raise NursePlanDatabaseError("Die gespeicherte Sitzung ist nicht mehr gültig.")
    rotated_token = _refresh_token_from_response(response)
    if not rotated_token:
        raise NursePlanDatabaseError("Die Sitzung konnte nicht erneuert werden.")
    return rotated_token


def current_session_refresh_token() -> str:
    """Liefert den aktuellen Token, damit eine automatische Rotation gespeichert wird."""
    if not is_authenticated():
        return ""
    session = get_supabase_client().auth.get_session()
    return str(getattr(session, "refresh_token", "") or "")


def clear_local_auth_state() -> None:
    """Entfernt ausschließlich lokale Authentifizierungsdaten dieser Streamlit-Sitzung."""
    for key in ["auth_user_id", "auth_email", "auth_name", "supabase_client"]:
        st.session_state.pop(key, None)


def sign_out_user() -> None:
    client = st.session_state.get("supabase_client")
    try:
        if client is not None and st.session_state.get("auth_user_id"):
            client.auth.sign_out()
    finally:
        clear_local_auth_state()


def is_authenticated() -> bool:
    return bool(
        st.session_state.get("supabase_client")
        and st.session_state.get("auth_user_id")
    )


def current_user_email() -> str:
    return str(st.session_state.get("auth_email", ""))


def _client() -> Client:
    if not is_authenticated():
        raise NursePlanDatabaseError("Bitte zuerst anmelden.")
    return get_supabase_client()


def _owner_id() -> str:
    owner_id = str(st.session_state.get("auth_user_id", ""))
    if not owner_id:
        raise NursePlanDatabaseError("Die Benutzer-Sitzung fehlt.")
    return owner_id


def _rows(response: Any) -> list[dict[str, Any]]:
    data = getattr(response, "data", None)
    if not data:
        return []
    if isinstance(data, dict):
        return [dict(data)]
    return [dict(row) for row in data]


def _new_id(response: Any, operation: str) -> int:
    rows = _rows(response)
    if not rows or "id" not in rows[0]:
        raise NursePlanDatabaseError(f"{operation} wurde nicht bestaetigt.")
    return int(rows[0]["id"])


def _json_value(value: Any, fallback: Any) -> Any:
    if isinstance(value, str):
        try:
            return json.loads(value)
        except (TypeError, json.JSONDecodeError):
            return fallback
    return value if value is not None else fallback


def _iso_date(value: str) -> str:
    text = str(value).strip()
    for date_format in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, date_format).date().isoformat()
        except ValueError:
            continue
    raise NursePlanDatabaseError(f"Ungültiges Datum: {text}")


def create_tables() -> None:
    """Die Tabellen werden einmalig mit supabase_schema.sql angelegt."""


def add_mitarbeitende(
    name: str,
    stellenanteil: str,
    wochenstunden: float,
    qualifikation: str,
    erfahrung: str,
    nachtdienst: str,
    monatssollstunden: float | None = None,
    schutzstatus: str = "Keiner",
    feste_dienstart: str = "Alle Dienste",
) -> int:
    monthly = (
        float(monatssollstunden)
        if monatssollstunden is not None
        else round(float(wochenstunden) * 52.0 / 12.0, 2)
    )
    response = (
        _client()
        .table("mitarbeitende")
        .insert(
            {
                "owner_id": _owner_id(),
                "name": name,
                "stellenanteil": stellenanteil,
                "wochenstunden": float(wochenstunden),
                "qualifikation": qualifikation,
                "erfahrung": erfahrung,
                "nachtdienst": nachtdienst,
                "monatssollstunden": monthly,
                "schutzstatus": schutzstatus,
                "feste_dienstart": feste_dienstart,
            }
        )
        .select("id")
        .execute()
    )
    return _new_id(response, "Das Speichern der Mitarbeitenden")


def update_mitarbeitende(
    mitarbeitende_id: int,
    name: str,
    stellenanteil: str,
    wochenstunden: float,
    qualifikation: str,
    erfahrung: str,
    nachtdienst: str,
    monatssollstunden: float | None = None,
    schutzstatus: str = "Keiner",
    feste_dienstart: str = "Alle Dienste",
) -> None:
    monthly = (
        float(monatssollstunden)
        if monatssollstunden is not None
        else round(float(wochenstunden) * 52.0 / 12.0, 2)
    )
    response = (
        _client()
        .table("mitarbeitende")
        .update(
            {
                "name": name,
                "stellenanteil": stellenanteil,
                "wochenstunden": float(wochenstunden),
                "qualifikation": qualifikation,
                "erfahrung": erfahrung,
                "nachtdienst": nachtdienst,
                "monatssollstunden": monthly,
                "schutzstatus": schutzstatus,
                "feste_dienstart": feste_dienstart,
            }
        )
        .eq("id", int(mitarbeitende_id))
        .select("id")
        .execute()
    )
    if not _rows(response):
        raise NursePlanDatabaseError("Die mitarbeitende Person wurde nicht gefunden.")


def get_mitarbeitende() -> list[dict[str, Any]]:
    response = (
        _client()
        .table("mitarbeitende")
        .select(
            "id,name,stellenanteil,wochenstunden,qualifikation,erfahrung,"
            "nachtdienst,monatssollstunden,schutzstatus,feste_dienstart"
        )
        .order("name")
        .execute()
    )
    return _rows(response)


def delete_mitarbeitende(mitarbeitende_id: int) -> None:
    (
        _client()
        .table("mitarbeitende")
        .delete()
        .eq("id", int(mitarbeitende_id))
        .execute()
    )


def save_dienstplan(
    name: str,
    startdatum: str,
    zeitraum: int,
    plan: list[dict[str, Any]],
    plan_id: int | None = None,
) -> int:
    payload = {
        "name": name,
        "startdatum": _iso_date(startdatum),
        "zeitraum": int(zeitraum),
        "plan_json": plan,
    }
    if plan_id is None:
        payload["owner_id"] = _owner_id()
        response = (
            _client().table("dienstplaene").insert(payload).select("id").execute()
        )
        return _new_id(response, "Der Dienstplan")

    response = (
        _client()
        .table("dienstplaene")
        .update(payload)
        .eq("id", int(plan_id))
        .select("id")
        .execute()
    )
    if not _rows(response):
        raise NursePlanDatabaseError("Der Dienstplan wurde nicht gefunden.")
    return int(plan_id)


def get_dienstplaene() -> list[dict[str, Any]]:
    response = (
        _client()
        .table("dienstplaene")
        .select("id,name,startdatum,zeitraum,erstellt_am,aktualisiert_am")
        .order("aktualisiert_am", desc=True)
        .execute()
    )
    return _rows(response)


def load_dienstplan(plan_id: int) -> dict[str, Any] | None:
    response = (
        _client()
        .table("dienstplaene")
        .select(
            "id,name,startdatum,zeitraum,plan_json,erstellt_am,aktualisiert_am"
        )
        .eq("id", int(plan_id))
        .limit(1)
        .execute()
    )
    rows = _rows(response)
    if not rows:
        return None
    result = rows[0]
    result["plan"] = _json_value(result.pop("plan_json", []), [])
    return result


def delete_dienstplan(plan_id: int) -> None:
    _client().table("dienstplaene").delete().eq("id", int(plan_id)).execute()


def get_einstellungen(defaults: dict[str, Any] | None = None) -> dict[str, Any]:
    result = dict(defaults or {})
    response = (
        _client()
        .table("einstellungen")
        .select("schluessel,wert_json")
        .execute()
    )
    for row in _rows(response):
        key = str(row.get("schluessel", ""))
        if key:
            result[key] = _json_value(row.get("wert_json"), result.get(key))
    return result


def save_einstellungen(values: dict[str, Any]) -> None:
    owner_id = _owner_id()
    rows = [
        {
            "owner_id": owner_id,
            "schluessel": str(key),
            "wert_json": value,
        }
        for key, value in values.items()
    ]
    if rows:
        (
            _client()
            .table("einstellungen")
            .upsert(rows, on_conflict="owner_id,schluessel")
            .execute()
        )


def add_feedback(
    kategorie: str,
    bewertung: int,
    betreff: str,
    nachricht: str,
    kontakt_erlaubt: bool,
    kontakt_email: str,
) -> int:
    allowed_categories = {"Idee", "Problem", "Frage", "Lob"}
    category = str(kategorie).strip()
    subject = str(betreff).strip()
    message = str(nachricht).strip()
    rating = int(bewertung)
    allow_contact = bool(kontakt_erlaubt)

    if category not in allowed_categories:
        raise NursePlanDatabaseError("Bitte eine gültige Feedback-Kategorie auswählen.")
    if not 1 <= rating <= 5:
        raise NursePlanDatabaseError("Die Bewertung muss zwischen 1 und 5 liegen.")
    if not 3 <= len(subject) <= 120:
        raise NursePlanDatabaseError("Der Betreff muss zwischen 3 und 120 Zeichen lang sein.")
    if not 10 <= len(message) <= 3000:
        raise NursePlanDatabaseError("Die Nachricht muss zwischen 10 und 3000 Zeichen lang sein.")

    response = (
        _client()
        .table("feedback")
        .insert(
            {
                "owner_id": _owner_id(),
                "kategorie": category,
                "bewertung": rating,
                "betreff": subject,
                "nachricht": message,
                "kontakt_erlaubt": allow_contact,
                "kontakt_email": (
                    str(kontakt_email).strip().lower() if allow_contact else ""
                ),
            }
        )
        .select("id")
        .execute()
    )
    return _new_id(response, "Das Feedback")


def get_feedback(limit: int = 20) -> list[dict[str, Any]]:
    """Liefert durch RLS und einen zusätzlichen Filter nur eigenes Feedback."""
    response = (
        _client()
        .table("feedback")
        .select(
            "id,kategorie,bewertung,betreff,nachricht,kontakt_erlaubt,"
            "kontakt_email,status,erstellt_am"
        )
        .eq("owner_id", _owner_id())
        .order("erstellt_am", desc=True)
        .limit(max(1, min(int(limit), 100)))
        .execute()
    )
    return _rows(response)


def add_vorgabe(
    mitarbeitende_id: int,
    code: str,
    startdatum: str,
    enddatum: str,
    bemerkung: str = "",
) -> int:
    response = (
        _client()
        .table("vorgaben")
        .insert(
            {
                "owner_id": _owner_id(),
                "mitarbeitende_id": int(mitarbeitende_id),
                "code": str(code),
                "startdatum": str(startdatum),
                "enddatum": str(enddatum),
                "bemerkung": str(bemerkung),
            }
        )
        .select("id")
        .execute()
    )
    return _new_id(response, "Die Abwesenheit")


def get_vorgaben() -> list[dict[str, Any]]:
    response = (
        _client()
        .table("vorgaben")
        .select(
            "id,mitarbeitende_id,code,startdatum,enddatum,bemerkung,erstellt_am"
        )
        .order("startdatum")
        .execute()
    )
    return _rows(response)


def delete_vorgabe(vorgabe_id: int) -> None:
    _client().table("vorgaben").delete().eq("id", int(vorgabe_id)).execute()


def save_vorlage(
    name: str,
    beschreibung: str,
    kategorie: str,
    daten: dict[str, Any],
    vorlage_id: int | None = None,
) -> int:
    payload = {
        "name": name.strip(),
        "beschreibung": beschreibung.strip(),
        "kategorie": kategorie.strip(),
        "daten_json": daten,
    }
    if vorlage_id is None:
        payload["owner_id"] = _owner_id()
        response = (
            _client().table("vorlagen").insert(payload).select("id").execute()
        )
        return _new_id(response, "Die Vorlage")

    response = (
        _client()
        .table("vorlagen")
        .update(payload)
        .eq("id", int(vorlage_id))
        .select("id")
        .execute()
    )
    if not _rows(response):
        raise NursePlanDatabaseError("Die Vorlage wurde nicht gefunden.")
    return int(vorlage_id)


def get_vorlagen() -> list[dict[str, Any]]:
    response = (
        _client()
        .table("vorlagen")
        .select("id,name,beschreibung,kategorie,erstellt_am,aktualisiert_am")
        .order("aktualisiert_am", desc=True)
        .execute()
    )
    return _rows(response)


def load_vorlage(vorlage_id: int) -> dict[str, Any] | None:
    response = (
        _client()
        .table("vorlagen")
        .select(
            "id,name,beschreibung,kategorie,daten_json,erstellt_am,aktualisiert_am"
        )
        .eq("id", int(vorlage_id))
        .limit(1)
        .execute()
    )
    rows = _rows(response)
    if not rows:
        return None
    result = rows[0]
    result["daten"] = _json_value(result.pop("daten_json", {}), {})
    return result


def delete_vorlage(vorlage_id: int) -> None:
    _client().table("vorlagen").delete().eq("id", int(vorlage_id)).execute()
