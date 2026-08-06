import calendar
import json
import re
from datetime import date, datetime, time, timedelta
from io import BytesIO
from uuid import uuid4
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import extra_streamlit_components as stx

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

from database import (
    SupabaseConfigurationError,
    add_feedback,
    add_mitarbeitende,
    add_vorgabe,
    create_tables,
    clear_local_auth_state,
    current_session_refresh_token,
    current_user_email,
    delete_dienstplan,
    delete_mitarbeitende,
    delete_vorgabe,
    get_dienstplaene,
    get_einstellungen,
    get_feedback,
    get_mitarbeitende,
    get_supabase_public_config,
    get_supabase_client,
    get_vorgaben,
    get_vorlagen,
    is_authenticated,
    load_dienstplan,
    load_vorlage,
    restore_user_session,
    request_password_reset,
    save_dienstplan,
    save_einstellungen,
    save_vorlage,
    delete_vorlage,
    sign_in_user,
    sign_out_user,
    sign_up_user,
    update_mitarbeitende,
)

st.set_page_config(page_title="NursePlan Pro", page_icon="🏥", layout="wide")

AUTH_COOKIE_NAME = "nurseplan_pro_session_v1"
AUTH_COOKIE_DAYS = 30

QUALIFIKATIONEN = [
    "Pflegefachkraft",
    "Praxisanleitung",
    "Stroke Nurse",
    "Schüler/in",
    "Pflegehelfer/in",
    "Stationshelfer/in",
    "Stationsleitung",
    "Stationssekretär/in",
]

FESTE_DIENSTARTEN = [
    "Alle Dienste",
    "Nur Frühdienst",
    "Nur Spätdienst",
    "Nur Nachtdienst",
]

FESTE_DIENSTART_CODES = {
    "Alle Dienste": {"F", "S", "N"},
    "Nur Frühdienst": {"F"},
    "Nur Spätdienst": {"S"},
    "Nur Nachtdienst": {"N"},
}

DIENST_CODES = ["", "F", "S", "N", "U", "W", "FB", "K", "Frei"]
WOCHENTAGE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
MONATSNAMEN = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember",
}

DEFAULT_SETTINGS = {
    "vollzeit_monatssoll": 167.0,
    "f_start": "06:00",
    "f_ende": "14:12",
    "f_stunden": 7.7,
    "min_f": 2,
    "s_start": "13:48",
    "s_ende": "22:00",
    "s_stunden": 7.7,
    "min_s": 2,
    "n_start": "21:48",
    "n_ende": "06:00",
    "n_stunden": 7.7,
    "min_n": 1,
    "min_fach_f": 1,
    "min_fach_s": 1,
    "min_fach_n": 1,
    "ruhezeit": 11.0,
    "max_arbeitstage": 6,
    "max_nachtdienste": 4,
    "max_wochenenden": 2,
    "keine_folgewochenenden": True,
    "max_plusstunden": 0.0,
    "regel_ruhezeit_aktiv": True,
    "regel_max_arbeitstage_aktiv": True,
    "regel_max_nachtdienste_aktiv": True,
    "regel_max_wochenenden_aktiv": True,
    "regel_sollstunden_aktiv": True,
    "regel_fachkraft_aktiv": True,
    "regel_nacht_zuerst_aktiv": True,
    "regel_nachtblock_aktiv": True,
    "regel_wochenende_zuerst_aktiv": True,
    "regel_wochenende_gleich_aktiv": True,
    "eigene_planungsregeln": [],
}

st.markdown(
    """
    <style>
    :root {
        --np-navy: #16324f;
        --np-blue: #2563eb;
        --np-blue-soft: #eaf2ff;
        --np-green: #0f9f6e;
        --np-orange: #d97706;
        --np-red: #dc2626;
        --np-text: #172033;
        --np-muted: #667085;
        --np-line: #e5e9f0;
        --np-bg: #f4f7fb;
        --np-card: #ffffff;
    }

    .stApp {
        background: var(--np-bg);
        color: var(--np-text);
        color-scheme: light;
    }
    header[data-testid="stHeader"] {
        background: rgba(244, 247, 251, .98) !important;
        border-bottom: 1px solid var(--np-line);
    }
    [data-testid="stMainBlockContainer"] {
        max-width: 1600px;
        padding: 1.25rem 1.6rem 3rem 1.6rem;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #102a43 0%, #173d5f 100%);
        border-right: 0;
    }
    section[data-testid="stSidebar"] * { color: #f8fafc !important; }
    section[data-testid="stSidebar"] [data-testid="stRadio"] label {
        padding: .6rem .75rem;
        border-radius: 10px;
        margin-bottom: .15rem;
    }
    section[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
        background: rgba(255,255,255,.09);
    }

    .np-brand {
        padding: .25rem 0 1rem 0;
        border-bottom: 1px solid rgba(255,255,255,.16);
        margin-bottom: .75rem;
    }
    .np-brand-name { font-size: 1.25rem; font-weight: 800; letter-spacing: -.02em; }
    .np-brand-subtitle { font-size: .82rem; opacity: .76; margin-top: .2rem; }

    .np-title {
        font-size: 2.15rem;
        line-height: 1.15;
        font-weight: 800;
        letter-spacing: -.035em;
        margin: 0;
        color: var(--np-text);
    }
    .np-subtitle {
        color: var(--np-muted);
        font-size: 1rem;
        margin: .35rem 0 1.3rem 0;
        max-width: 900px;
    }
    .np-section-title {
        font-size: 1.12rem;
        font-weight: 750;
        margin: .2rem 0 .2rem 0;
    }
    .np-help { color: var(--np-muted); font-size: .9rem; margin-bottom: .8rem; }

    .np-card {
        background: var(--np-card);
        border: 1px solid var(--np-line);
        border-radius: 16px;
        padding: 1.05rem 1.15rem;
        margin-bottom: 1rem;
        box-shadow: 0 4px 16px rgba(15, 23, 42, .035);
    }
    .np-kpi {
        background: var(--np-card);
        border: 1px solid var(--np-line);
        border-radius: 15px;
        padding: 1rem 1.05rem;
        min-height: 112px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, .035);
    }
    .np-kpi-label { color: var(--np-muted); font-size: .82rem; font-weight: 650; }
    .np-kpi-value { color: var(--np-text); font-size: 1.8rem; font-weight: 800; margin-top: .15rem; }
    .np-kpi-note { color: var(--np-muted); font-size: .78rem; margin-top: .12rem; }

    .np-percent {
        background: var(--np-blue-soft);
        border: 1px solid #cfddff;
        border-radius: 12px;
        padding: .7rem .9rem;
        margin: .35rem 0 .8rem 0;
    }
    .np-percent-label { color: var(--np-muted); font-size: .8rem; }
    .np-percent-value { color: #1d4ed8; font-size: 1.35rem; font-weight: 800; }

    .np-table-wrap {
        overflow-x: auto;
        border: 1px solid var(--np-line);
        border-radius: 14px;
        background: white;
    }
    .np-table { width: 100%; border-collapse: collapse; min-width: 780px; }
    .np-table th {
        text-align: left;
        background: #f8fafc;
        color: #475467;
        font-size: .78rem;
        text-transform: uppercase;
        letter-spacing: .035em;
        padding: .72rem .8rem;
        border-bottom: 1px solid var(--np-line);
    }
    .np-table td {
        padding: .72rem .8rem;
        border-bottom: 1px solid #eef1f5;
        font-size: .9rem;
        vertical-align: middle;
    }
    .np-table tr:last-child td { border-bottom: 0; }
    .np-table tr:hover td { background: #fbfdff; }
    .np-name { font-weight: 750; color: var(--np-text); }
    .np-badge {
        display: inline-block;
        border-radius: 999px;
        padding: .2rem .55rem;
        font-size: .75rem;
        font-weight: 700;
        white-space: nowrap;
    }
    .np-badge-blue { background: #eaf2ff; color: #1d4ed8; }
    .np-badge-green { background: #e8f8f2; color: #087a54; }
    .np-badge-gray { background: #f1f3f6; color: #475467; }
    .np-badge-orange { background: #fff4df; color: #a85b00; }
    .np-badge-red { background: #feecec; color: #b42318; }

    .np-template-card {
        background: white;
        border: 1px solid var(--np-line);
        border-radius: 15px;
        padding: 1rem;
        min-height: 195px;
        margin-bottom: .45rem;
    }
    .np-template-title { font-weight: 800; font-size: 1.02rem; margin-bottom: .35rem; }
    .np-template-text { color: var(--np-muted); font-size: .86rem; line-height: 1.45; }

    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
        background: white;
        border: 1px solid var(--np-line);
        border-radius: 14px;
        overflow: hidden;
    }
    div[data-testid="stForm"] {
        background: white;
        border: 1px solid var(--np-line);
        border-radius: 15px;
        padding: 1.15rem;
    }
    [data-testid="stMain"] [data-testid="stWidgetLabel"] p,
    [data-testid="stMain"] div[data-testid="stTextInput"] label p,
    [data-testid="stMain"] div[data-testid="stNumberInput"] label p,
    [data-testid="stMain"] div[data-testid="stSelectbox"] label p,
    [data-testid="stMain"] div[data-testid="stDateInput"] label p,
    [data-testid="stMain"] div[data-testid="stTextArea"] label p {
        color: #344054 !important;
        font-size: .88rem !important;
        font-weight: 700 !important;
    }
    [data-testid="stMain"] div[data-baseweb="input"],
    [data-testid="stMain"] div[data-baseweb="base-input"],
    [data-testid="stMain"] div[data-baseweb="input"] > div {
        background: #ffffff !important;
        border-color: #cfd6e1 !important;
        color: #172033 !important;
    }
    [data-testid="stMain"] div[data-baseweb="input"] {
        border: 1px solid #cfd6e1 !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 2px rgba(16, 24, 40, .04);
    }
    [data-testid="stMain"] div[data-baseweb="input"]:focus-within {
        border-color: var(--np-blue) !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, .12) !important;
    }
    [data-testid="stMain"] div[data-baseweb="input"] input {
        color: #172033 !important;
        -webkit-text-fill-color: #172033 !important;
        caret-color: var(--np-blue) !important;
    }
    [data-testid="stMain"] div[data-baseweb="input"] input::placeholder {
        color: #98a2b3 !important;
        -webkit-text-fill-color: #98a2b3 !important;
    }
    [data-testid="stMain"] div[data-baseweb="input"] svg {
        color: #667085 !important;
        fill: #667085 !important;
    }
    [data-testid="stMain"] div[data-baseweb="input"] button {
        background: transparent !important;
        color: #667085 !important;
    }
    div[data-testid="stExpander"] {
        border: 1px solid var(--np-line);
        border-radius: 12px;
        background: white;
    }
    .stButton button, .stDownloadButton button, .stFormSubmitButton button {
        border-radius: 10px;
        font-weight: 700;
        min-height: 2.55rem;
    }
    button[kind="primary"] {
        background: var(--np-blue) !important;
        border-color: var(--np-blue) !important;
        color: #ffffff !important;
        box-shadow: 0 3px 8px rgba(37, 99, 235, .18);
    }
    button[kind="primary"] p { color: #ffffff !important; }
    button[kind="primary"]:hover {
        background: #1d4ed8 !important;
        border-color: #1d4ed8 !important;
    }
    .stTabs [data-baseweb="tab-list"] { gap: .35rem; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 9px 9px 0 0;
        padding: .55rem .9rem;
        font-weight: 700;
        color: #667085 !important;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        color: #1d4ed8 !important;
    }
    .stTabs [data-baseweb="tab"] p { color: inherit !important; }
    .stTabs [data-baseweb="tab-highlight"] {
        background-color: var(--np-blue) !important;
    }
    hr { border-color: var(--np-line); }
    </style>
    """,
    unsafe_allow_html=True,
)


auth_cookie_manager = stx.CookieManager(key="nurseplan_auth_cookie_manager")


def auth_cookie_is_secure() -> bool:
    """Secure-Cookies auf HTTPS; localhost bleibt für die Entwicklung nutzbar."""
    try:
        return str(st.context.url).casefold().startswith("https://")
    except Exception:
        return False


def saved_auth_cookie() -> str:
    return str(auth_cookie_manager.get(AUTH_COOKIE_NAME) or "").strip()


def save_auth_cookie(refresh_token: str) -> None:
    token = str(refresh_token or "").strip()
    if not token or saved_auth_cookie() == token:
        return
    auth_cookie_manager.set(
        AUTH_COOKIE_NAME,
        token,
        key="save_nurseplan_auth_cookie",
        path="/",
        expires_at=datetime.now() + timedelta(days=AUTH_COOKIE_DAYS),
        max_age=AUTH_COOKIE_DAYS * 24 * 60 * 60,
        secure=auth_cookie_is_secure(),
        same_site="strict",
    )


def delete_auth_cookie() -> None:
    if not saved_auth_cookie():
        return
    auth_cookie_manager.delete(
        AUTH_COOKIE_NAME,
        key="delete_nurseplan_auth_cookie",
    )


def auth_error_text(error: Exception, action: str) -> str:
    message = str(error).casefold()
    if "invalid login credentials" in message:
        return "E-Mail oder Passwort ist nicht richtig."
    if "email not confirmed" in message:
        return "Bitte bestätige zuerst deine E-Mail-Adresse."
    if "already registered" in message:
        return "Für diese E-Mail-Adresse besteht bereits ein Konto."
    if "rate limit" in message or "too many" in message:
        return "Zu viele Versuche. Bitte warte kurz und versuche es erneut."
    if "password" in message and ("least" in message or "short" in message):
        return "Das Passwort ist zu kurz. Verwende mindestens 8 Zeichen."
    if action == "login":
        return "Die Anmeldung ist gerade nicht möglich. Bitte versuche es erneut."
    if action == "password_reset":
        return "Die E-Mail konnte gerade nicht gesendet werden. Bitte versuche es später erneut."
    return "Die Registrierung ist gerade nicht möglich. Bitte versuche es erneut."


def password_reset_redirect_url() -> str:
    app_url = configured_app_url()
    return f"{app_url.rstrip('/')}/?password_recovery=1"


def configured_app_url() -> str:
    """Liefert die öffentliche App-Adresse ohne Pfad oder abschließenden Slash."""
    try:
        configured_url = str(st.secrets.get("APP_URL", "") or "").strip()
    except Exception:
        configured_url = ""
    return (configured_url or "http://localhost:8501").rstrip("/")


def google_oauth_redirect_url() -> str:
    return f"{configured_app_url()}/?oauth_callback=google"


def google_oauth_url() -> str:
    """Erstellt die Supabase-Adresse für Anmeldung und Registrierung mit Google."""
    response = get_supabase_client().auth.sign_in_with_oauth(
        {
            "provider": "google",
            "options": {"redirect_to": google_oauth_redirect_url()},
        }
    )
    url = str(getattr(response, "url", "") or "").strip()
    if not url:
        raise RuntimeError("Die Google-Anmeldung konnte nicht gestartet werden.")
    return url


def show_google_oauth_callback() -> None:
    """Übernimmt nach Google OAuth den Supabase-Refresh-Token in das App-Cookie."""
    secure_cookie = "true" if auth_cookie_is_secure() else "false"
    callback_html = r"""
<!doctype html>
<html lang="de">
<head><meta charset="utf-8"></head>
<body>
<script>
(() => {
  const fragment = new URLSearchParams(window.parent.location.hash.slice(1));
  const refreshToken = fragment.get("refresh_token") || "";
  const secureCookie = __SECURE_COOKIE__;

  if (!refreshToken) {
    window.parent.history.replaceState({}, "", `${window.parent.location.pathname}?oauth_error=1`);
    window.parent.location.reload();
    return;
  }

  let cookie = "__COOKIE_NAME__=" + encodeURIComponent(refreshToken)
    + "; Path=/; Max-Age=__COOKIE_MAX_AGE__; SameSite=Lax";
  if (secureCookie) cookie += "; Secure";
  window.parent.document.cookie = cookie;
  window.parent.history.replaceState({}, "", `${window.parent.location.pathname}?oauth_complete=1`);
  window.parent.location.reload();
})();
</script>
</body>
</html>
"""
    callback_html = callback_html.replace("__SECURE_COOKIE__", secure_cookie)
    callback_html = callback_html.replace("__COOKIE_NAME__", AUTH_COOKIE_NAME)
    callback_html = callback_html.replace(
        "__COOKIE_MAX_AGE__", str(AUTH_COOKIE_DAYS * 24 * 60 * 60)
    )

    left, middle, right = st.columns([1, 1.25, 1])
    with middle:
        st.markdown(
            """
            <div class="np-card" style="text-align:center;margin-top:3rem;">
                <div style="font-size:2.2rem;">🏥</div>
                <div style="font-size:1.55rem;font-weight:800;">Google-Anmeldung wird abgeschlossen</div>
                <div style="color:#667085;margin-top:.35rem;">Einen Moment bitte …</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        components.html(callback_html, height=1, scrolling=False)


def show_password_recovery_page() -> None:
    """Zeigt das neue Passwortfeld; das Passwort geht direkt an Supabase Auth."""
    supabase_url, publishable_key = get_supabase_public_config()
    recovery_html = r"""
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 8px; background: transparent; color: #172033;
    font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  }
  .card {
    background: #fff; border: 1px solid #e5e9f0; border-radius: 16px;
    padding: 24px; box-shadow: 0 12px 32px rgba(16,42,67,.08);
  }
  h2 { margin: 0 0 6px; font-size: 22px; }
  .subtitle { margin: 0 0 20px; color: #667085; font-size: 14px; line-height: 1.45; }
  label { display: block; margin: 14px 0 6px; font-size: 14px; font-weight: 700; }
  input {
    width: 100%; border: 1px solid #cfd6e2; border-radius: 10px;
    padding: 12px 13px; font-size: 15px; background: #fff; color: #172033;
  }
  input:focus { outline: 3px solid #dbeafe; border-color: #2563eb; }
  button {
    width: 100%; margin-top: 18px; border: 0; border-radius: 10px;
    padding: 12px 16px; background: #2563eb; color: white;
    font-size: 15px; font-weight: 750; cursor: pointer;
  }
  button:disabled { opacity: .65; cursor: wait; }
  .message { display: none; margin-top: 14px; border-radius: 10px; padding: 11px 12px; font-size: 14px; }
  .error { display: block; background: #fef2f2; color: #b42318; }
  .success { display: block; background: #ecfdf3; color: #067647; }
</style>
</head>
<body>
  <div class="card">
    <h2>Neues Passwort festlegen</h2>
    <p class="subtitle">Wähle ein sicheres Passwort mit mindestens 8 Zeichen.</p>
    <form id="passwordForm">
      <label for="password">Neues Passwort</label>
      <input id="password" type="password" minlength="8" autocomplete="new-password" required>
      <label for="passwordRepeat">Passwort wiederholen</label>
      <input id="passwordRepeat" type="password" minlength="8" autocomplete="new-password" required>
      <button id="submitButton" type="submit">Passwort speichern</button>
    </form>
    <div id="message" class="message" role="status"></div>
  </div>
<script>
(() => {
  const supabaseUrl = __SUPABASE_URL__;
  const publishableKey = __PUBLISHABLE_KEY__;
  const form = document.getElementById("passwordForm");
  const message = document.getElementById("message");
  const submitButton = document.getElementById("submitButton");

  let accessToken = "";
  try {
    const fragment = new URLSearchParams(window.parent.location.hash.slice(1));
    accessToken = fragment.get("access_token") || "";
    if (fragment.get("type") !== "recovery") accessToken = "";
  } catch (_) {
    accessToken = "";
  }

  if (!accessToken) {
    form.style.display = "none";
    message.textContent = "Der Link ist ungültig oder abgelaufen. Bitte fordere einen neuen Link an.";
    message.className = "message error";
    return;
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    message.className = "message";
    const password = document.getElementById("password").value;
    const repeated = document.getElementById("passwordRepeat").value;
    if (password.length < 8) {
      message.textContent = "Das Passwort muss mindestens 8 Zeichen haben.";
      message.className = "message error";
      return;
    }
    if (password !== repeated) {
      message.textContent = "Die beiden Passwörter stimmen nicht überein.";
      message.className = "message error";
      return;
    }

    submitButton.disabled = true;
    submitButton.textContent = "Wird gespeichert …";
    try {
      const response = await fetch(`${supabaseUrl}/auth/v1/user`, {
        method: "PUT",
        headers: {
          "apikey": publishableKey,
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ password })
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.msg || data.message || "Passwort konnte nicht geändert werden.");
      form.style.display = "none";
      message.textContent = "Dein Passwort wurde geändert. Klicke jetzt unten auf „Zurück zur Anmeldung“.";
      message.className = "message success";
    } catch (_) {
      message.textContent = "Das Passwort konnte nicht geändert werden. Bitte fordere einen neuen Link an.";
      message.className = "message error";
      submitButton.disabled = false;
      submitButton.textContent = "Passwort speichern";
    }
  });
})();
</script>
</body>
</html>
"""
    recovery_html = recovery_html.replace("__SUPABASE_URL__", json.dumps(supabase_url))
    recovery_html = recovery_html.replace("__PUBLISHABLE_KEY__", json.dumps(publishable_key))

    left, middle, right = st.columns([1, 1.25, 1])
    with middle:
        st.markdown(
            """
            <div class="np-card" style="text-align:center;margin-top:3rem;">
                <div style="font-size:2.2rem;">🏥</div>
                <div style="font-size:1.65rem;font-weight:800;">NursePlan Pro</div>
                <div style="color:#667085;margin-top:.25rem;">Sicheres Passwort festlegen</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        components.html(recovery_html, height=390, scrolling=False)
        if st.button(
            "← Zurück zur Anmeldung",
            type="secondary",
            width="stretch",
            key="password_recovery_back_to_login",
        ):
            delete_auth_cookie()
            clear_local_auth_state()
            st.session_state.auth_logout_requested = True
            st.query_params.clear()
            st.rerun()


def show_auth_page() -> None:
    left, middle, right = st.columns([1, 1.25, 1])
    with middle:
        st.markdown(
            """
            <div class="np-card" style="text-align:center;margin-top:3rem;">
                <div style="font-size:2.2rem;">🏥</div>
                <div style="font-size:1.65rem;font-weight:800;">Willkommen bei NursePlan Pro</div>
                <div style="color:#667085;margin-top:.25rem;">Dienstplanung sicher, übersichtlich und überall verfügbar</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.session_state.pop("password_reset_success", False):
            st.success("Dein Passwort wurde geändert. Du kannst dich jetzt anmelden.")
        if st.session_state.pop("google_oauth_error", False):
            st.error(
                "Die Google-Anmeldung wurde abgebrochen oder konnte nicht abgeschlossen werden."
            )

        if st.session_state.get("auth_show_password_reset"):
            st.markdown("#### Passwort zurücksetzen")
            st.caption("Wir senden dir einen sicheren Link an deine E-Mail-Adresse.")
            if st.session_state.get("auth_reset_email_sent"):
                st.success(
                    "Wenn für diese Adresse ein Konto besteht, wurde eine E-Mail gesendet. "
                    "Bitte prüfe auch den Spam-Ordner."
                )
            else:
                with st.form("password_reset_request_form"):
                    reset_email = st.text_input(
                        "E-Mail-Adresse",
                        key="password_reset_email",
                        autocomplete="email",
                    )
                    submit_reset = st.form_submit_button(
                        "Link senden", type="primary", width="stretch"
                    )
                if submit_reset:
                    if not reset_email.strip() or "@" not in reset_email:
                        st.warning("Bitte gib eine gültige E-Mail-Adresse ein.")
                    else:
                        try:
                            request_password_reset(
                                reset_email,
                                password_reset_redirect_url(),
                            )
                            st.session_state.auth_reset_email_sent = True
                            st.rerun()
                        except Exception as exc:
                            st.error(auth_error_text(exc, "password_reset"))

            if st.button("← Zurück zur Anmeldung", width="stretch"):
                st.session_state.auth_show_password_reset = False
                st.session_state.pop("auth_reset_email_sent", None)
                st.rerun()

            st.caption(
                "Der Link kann nur einmal verwendet werden und läuft nach kurzer Zeit ab."
            )
            return

        try:
            google_url = google_oauth_url()
        except Exception:
            google_url = ""

        if google_url:
            safe_google_url = escape(google_url, {'"': "&quot;"})
            st.markdown(
                f"""
                <a href="{safe_google_url}" target="_top" style="
                    display:flex;align-items:center;justify-content:center;gap:.7rem;
                    width:100%;padding:.78rem 1rem;border:1px solid #cfd6e2;
                    border-radius:10px;background:#fff;color:#172033;
                    text-decoration:none;font-weight:750;box-shadow:0 2px 5px rgba(16,42,67,.05);">
                    <span style="font-size:1.12rem;font-weight:900;color:#4285f4;">G</span>
                    Mit Google anmelden
                </a>
                <div style="display:flex;align-items:center;gap:.75rem;margin:1rem 0;color:#98a2b3;">
                    <div style="height:1px;background:#e5e9f0;flex:1;"></div>
                    <span style="font-size:.82rem;">oder mit E-Mail</span>
                    <div style="height:1px;background:#e5e9f0;flex:1;"></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.warning("Die Google-Anmeldung ist vorübergehend nicht verfügbar.")

        login_tab, register_tab = st.tabs(["Anmelden", "Konto erstellen"])

        with login_tab:
            st.markdown("#### Sicher anmelden")
            st.caption("Melde dich mit deinem persönlichen NursePlan-Pro-Konto an.")
            with st.form("login_form"):
                email = st.text_input("E-Mail-Adresse", key="login_email")
                password = st.text_input("Passwort", type="password", key="login_password")
                remember_login = st.checkbox(
                    "Auf diesem Gerät angemeldet bleiben",
                    value=True,
                    help="Nur auf einem persönlichen oder geschützten Gerät verwenden.",
                )
                submit_login = st.form_submit_button(
                    "Anmelden", type="primary", width="stretch"
                )
            if submit_login:
                if not email.strip() or not password:
                    st.warning("Bitte E-Mail-Adresse und Passwort eingeben.")
                else:
                    try:
                        refresh_token = sign_in_user(email, password)
                        st.session_state.pop("auth_logout_requested", None)
                        st.session_state.pop("auth_restore_failed", None)
                        st.session_state.remember_login = bool(remember_login)
                        if remember_login:
                            save_auth_cookie(refresh_token)
                        else:
                            delete_auth_cookie()
                        st.rerun()
                    except Exception as exc:
                        st.error(auth_error_text(exc, "login"))

            if st.button("Passwort vergessen?", width="stretch"):
                st.session_state.auth_show_password_reset = True
                st.session_state.pop("auth_reset_email_sent", None)
                st.rerun()

        with register_tab:
            st.markdown("#### Persönliches Konto erstellen")
            st.caption("Deine Dienstpläne und Teamdaten bleiben deinem Konto zugeordnet.")
            with st.form("register_form"):
                full_name = st.text_input("Vollständiger Name", key="register_name")
                register_email = st.text_input(
                    "E-Mail-Adresse", key="register_email"
                )
                register_password = st.text_input(
                    "Passwort (mindestens 8 Zeichen)",
                    type="password",
                    key="register_password",
                )
                password_repeat = st.text_input(
                    "Passwort bestätigen",
                    type="password",
                    key="register_password_repeat",
                )
                remember_registration = st.checkbox(
                    "Auf diesem Gerät angemeldet bleiben",
                    value=True,
                    key="register_remember_login",
                    help="Nur auf einem persönlichen oder geschützten Gerät verwenden.",
                )
                submit_registration = st.form_submit_button(
                    "Konto erstellen", type="primary", width="stretch"
                )
            if submit_registration:
                if not full_name.strip() or not register_email.strip():
                    st.warning("Bitte Name und E-Mail-Adresse eingeben.")
                elif len(register_password) < 8:
                    st.warning("Das Passwort muss mindestens 8 Zeichen haben.")
                elif register_password != password_repeat:
                    st.warning("Die beiden Passwörter stimmen nicht überein.")
                else:
                    try:
                        refresh_token = sign_up_user(
                            full_name, register_email, register_password
                        )
                        if refresh_token:
                            st.session_state.pop("auth_logout_requested", None)
                            st.session_state.pop("auth_restore_failed", None)
                            st.session_state.remember_login = bool(remember_registration)
                            if remember_registration:
                                save_auth_cookie(refresh_token)
                            else:
                                delete_auth_cookie()
                            st.rerun()
                        else:
                            st.success(
                                "Das Konto wurde angelegt. Öffne jetzt die "
                                "Bestätigungs-E-Mail und melde dich danach an."
                            )
                    except Exception as exc:
                        st.error(auth_error_text(exc, "registration"))

        st.caption(
            "🔒 Kontogeschützt: Mitarbeitende, Abwesenheiten und Dienstpläne "
            "sind nur im angemeldeten Konto verfügbar."
        )


try:
    get_supabase_client()
except SupabaseConfigurationError as exc:
    st.error("Supabase ist noch nicht mit der App verbunden.")
    st.write(str(exc))
    st.code(
        'SUPABASE_URL = "https://DEIN-PROJEKT.supabase.co"\n'
        'SUPABASE_PUBLISHABLE_KEY = "DEIN-PUBLISHABLE-KEY"',
        language="toml",
    )
    st.caption(
        "Diese zwei Werte gehören lokal in .streamlit/secrets.toml. "
        "Niemals den Secret- oder service_role-Schlüssel verwenden."
    )
    st.stop()

# Nach der Rückkehr von Google liegt die Supabase-Sitzung kurz im URL-Fragment.
# Ein kleines Browser-Hilfsfenster speichert nur den Refresh-Token im vorhandenen
# NursePlan-Cookie und entfernt anschließend sämtliche Token aus der Adresse.
if str(st.query_params.get("oauth_callback", "")).casefold() == "google":
    show_google_oauth_callback()
    st.stop()

if str(st.query_params.get("oauth_complete", "")).casefold() == "1":
    st.session_state.pop("auth_logout_requested", None)
    st.session_state.pop("auth_restore_failed", None)
    st.query_params.clear()

if str(st.query_params.get("oauth_error", "")).casefold() == "1":
    st.session_state.google_oauth_error = True
    st.query_params.clear()

# Nach einer erfolgreichen Änderung kehrt der Browser ohne Sitzung zur Anmeldung zurück.
if str(st.query_params.get("password_reset", "")).casefold() == "success":
    delete_auth_cookie()
    clear_local_auth_state()
    st.session_state.auth_logout_requested = True
    st.session_state.password_reset_success = True
    st.query_params.clear()

# Der Supabase-Link enthält die kurzlebige Wiederherstellungs-Sitzung im URL-Fragment.
# Das Passwortformular verarbeitet sie direkt im Browser, damit kein Passwort die App erreicht.
if str(st.query_params.get("password_recovery", "")).casefold() == "1":
    show_password_recovery_page()
    st.stop()

# Eine gespeicherte Supabase-Sitzung wird nach Browser-Refresh erneuert.
# Refresh-Tokens werden dabei rotiert und sofort wieder im Cookie aktualisiert.
if (
    not is_authenticated()
    and not st.session_state.get("auth_logout_requested")
    and not st.session_state.get("auth_restore_failed")
):
    stored_refresh_token = saved_auth_cookie()
    if stored_refresh_token:
        try:
            rotated_refresh_token = restore_user_session(stored_refresh_token)
            st.session_state.remember_login = True
            save_auth_cookie(rotated_refresh_token)
        except Exception:
            clear_local_auth_state()
            st.session_state.auth_restore_failed = True

if is_authenticated() and st.session_state.get("remember_login"):
    current_refresh_token = current_session_refresh_token()
    if current_refresh_token:
        save_auth_cookie(current_refresh_token)

if not is_authenticated():
    show_auth_page()
    st.stop()

create_tables()

if "settings" not in st.session_state:
    raw_settings = get_einstellungen(DEFAULT_SETTINGS)
    # Werte aus älteren Versionen übernehmen.
    if "n_start" not in raw_settings and "na_start" in raw_settings:
        raw_settings["n_start"] = raw_settings.get("na_start", "21:48")
        raw_settings["n_ende"] = raw_settings.get("na_ende", "06:00")
        raw_settings["n_stunden"] = raw_settings.get("na_stunden", 7.7)
        raw_settings["min_n"] = raw_settings.get("min_na", 1)
    st.session_state.settings = {**DEFAULT_SETTINGS, **raw_settings}

if "dienstplan_vorlage" not in st.session_state:
    st.session_state.dienstplan_vorlage = []
if "dienstplan_editor_version" not in st.session_state:
    st.session_state.dienstplan_editor_version = 0
if "aktiver_dienstplan_id" not in st.session_state:
    st.session_state.aktiver_dienstplan_id = None
if "dienstplan_name" not in st.session_state:
    st.session_state.dienstplan_name = "Neuer Dienstplan"
if "auto_plan_result" not in st.session_state:
    st.session_state.auto_plan_result = None


def page_header(title: str, subtitle: str) -> None:
    st.markdown(f'<p class="np-title">{escape(title)}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="np-subtitle">{escape(subtitle)}</p>', unsafe_allow_html=True)


def parse_clock(value: str) -> time:
    try:
        return datetime.strptime(str(value), "%H:%M").time()
    except (TypeError, ValueError):
        return time(0, 0)


def format_clock(value: time) -> str:
    return value.strftime("%H:%M")


def percentage_from_hours(monthly_hours: float) -> float:
    basis = float(st.session_state.settings.get("vollzeit_monatssoll", 167.0))
    if basis <= 0:
        return 0.0
    return round(float(monthly_hours) / basis * 100.0, 1)


def percentage_text(monthly_hours: float) -> str:
    value = percentage_from_hours(monthly_hours)
    if value.is_integer():
        return f"{int(value)} %"
    return f"{value:.1f} %".replace(".", ",")


def weekly_from_month(monthly_hours: float) -> float:
    return round(float(monthly_hours) * 12.0 / 52.0, 2)


def stored_percentage(monthly_hours: float) -> str:
    return f"{percentage_from_hours(monthly_hours):.1f}%"


def monthly_target(person: dict) -> float:
    try:
        return round(float(person.get("monatssollstunden", 0.0)), 2)
    except (TypeError, ValueError):
        return 0.0


def day_column(day_value: date) -> str:
    return f"{day_value.day:02d} {WOCHENTAGE[day_value.weekday()]}"


def split_names(value) -> list[str]:
    if value is None or pd.isna(value):
        return []
    return [x.strip() for x in re.split(r"[,;\n]+", str(value)) if x.strip()]


def join_names(values) -> str:
    result = []
    seen = set()
    for value in values:
        name = str(value).strip()
        key = name.casefold()
        if name and key not in seen:
            result.append(name)
            seen.add(key)
    return ", ".join(result)


def parse_date(value):
    try:
        return datetime.strptime(str(value), "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


def weekend_anchor(day_value: date | None) -> date | None:
    """Gibt den Samstag eines Wochenendes zurück; Samstag und Sonntag gehören zusammen."""
    if day_value is None:
        return None
    if day_value.weekday() == 5:
        return day_value
    if day_value.weekday() == 6:
        return day_value - timedelta(days=1)
    return None


def worked_weekends(values) -> set[date]:
    return {anchor for value in values if (anchor := weekend_anchor(value)) is not None}


def has_consecutive_weekends(values: set[date]) -> bool:
    ordered = sorted(values)
    return any(current == previous + timedelta(days=7) for previous, current in zip(ordered, ordered[1:]))


def regel_aktiv(settings: dict, key: str, default: bool = True) -> bool:
    return bool(settings.get(key, default))


def eigene_planungsregeln(settings: dict, regeltyp: str | None = None) -> list[dict]:
    rules = settings.get("eigene_planungsregeln", [])
    if not isinstance(rules, list):
        return []
    result = [dict(rule) for rule in rules if isinstance(rule, dict)]
    if regeltyp is not None:
        result = [rule for rule in result if str(rule.get("typ")) == regeltyp]
    return result


def regel_gilt_fuer_person(rule: dict, person: dict) -> bool:
    try:
        person_id = int(rule.get("mitarbeitende_id", 0))
        return person_id == 0 or person_id == int(person.get("id", -1))
    except (TypeError, ValueError):
        return False


def regel_zahl(rule: dict, key: str, default: int) -> int:
    try:
        return int(rule.get(key, default))
    except (TypeError, ValueError):
        return int(default)


def mitarbeitende_name(employee_id, employees) -> str:
    try:
        target = int(employee_id)
    except (TypeError, ValueError):
        return "Unbekannt"
    if target == 0:
        return "Alle Mitarbeitenden"
    for person in employees:
        if int(person.get("id", -1)) == target:
            return str(person.get("name", "Unbekannt"))
    return "Gelöschte Person"


def regel_beschreibung(rule: dict, employees) -> str:
    shift_names = {"F": "Frühdienst", "S": "Spätdienst", "N": "Nachtdienst", "ALL": "allen Diensten"}
    rule_type = str(rule.get("typ", ""))
    if rule_type == "nicht_gemeinsam":
        first = mitarbeitende_name(rule.get("mitarbeitende_a_id"), employees)
        second = mitarbeitende_name(rule.get("mitarbeitende_b_id"), employees)
        code = str(rule.get("dienst", "ALL"))
        if code == "ALL":
            return f"{first} und {second} dürfen nicht gemeinsam in derselben Schicht arbeiten."
        shift = shift_names.get(code, "gewählten Dienst")
        return f"{first} und {second} dürfen nicht gemeinsam im {shift} eingeplant werden."
    if rule_type == "frei_nach_dienst":
        person = mitarbeitende_name(rule.get("mitarbeitende_id", 0), employees)
        shift = shift_names.get(str(rule.get("dienst", "N")), "Dienst")
        days = regel_zahl(rule, "freie_tage", 1)
        return f"{person}: Nach einem {shift} folgen mindestens {days} freie Kalendertage."
    if rule_type == "max_dienstfolge":
        person = mitarbeitende_name(rule.get("mitarbeitende_id", 0), employees)
        shift = shift_names.get(str(rule.get("dienst", "N")), "Dienst")
        maximum = regel_zahl(rule, "maximum", 1)
        return f"{person}: Höchstens {maximum} {shift}-Dienste hintereinander."
    return "Unbekannte Regelart"


def normalise_plan(plan):
    result = []
    for index, entry in enumerate(plan or [], start=1):
        row = dict(entry)
        row.setdefault("Tag", index)
        row.setdefault("Datum", "")
        row.setdefault("Frühdienst", "")
        row.setdefault("Spätdienst", "")
        row.setdefault("Nachtdienst", "")
        row.setdefault("Urlaub", "")
        row.setdefault("Wunschfrei", "")
        row.setdefault("Fortbildung", "")
        row.setdefault("Krank", "")
        row.setdefault("Frei", "")
        row.setdefault("Status", "")
        result.append(row)
    return result

def create_empty_month(year: int, month: int):
    return [
        {
            "Tag": day,
            "Datum": date(year, month, day).strftime("%d.%m.%Y"),
            "Frühdienst": "",
            "Spätdienst": "",
            "Nachtdienst": "",
            "Urlaub": "",
            "Wunschfrei": "",
            "Fortbildung": "",
            "Krank": "",
            "Frei": "",
            "Status": "",
        }
        for day in range(1, calendar.monthrange(year, month)[1] + 1)
    ]


FACHKRAFT_QUALIFIKATIONEN = {
    "Pflegefachkraft",
    "Praxisanleitung",
    "Stroke Nurse",
    "Stationsleitung",
}

VORGABE_LABELS = {
    "U": "Urlaub",
    "W": "Wunschfrei",
    "FB": "Fortbildung",
    "K": "Krank",
}

# Alte Versionen erlaubten F/S/N/Frei als datumsbezogene Vorgabe.
# Diese Einträge werden nicht mehr automatisch angewendet, können aber in der
# Übersicht noch erkannt und gelöscht werden.
LEGACY_VORGABE_LABELS = {
    "F": "Alte Vorgabe: Frühdienst (wird ignoriert)",
    "S": "Alte Vorgabe: Spätdienst (wird ignoriert)",
    "N": "Alte Vorgabe: Nachtdienst (wird ignoriert)",
    "Frei": "Alte Vorgabe: fest frei (wird ignoriert)",
}

CODE_TO_COLUMN = {
    "F": "Frühdienst",
    "S": "Spätdienst",
    "N": "Nachtdienst",
    "U": "Urlaub",
    "W": "Wunschfrei",
    "FB": "Fortbildung",
    "K": "Krank",
    "Frei": "Frei",
}


def is_fachkraft(person: dict) -> bool:
    return str(person.get("qualifikation", "")) in FACHKRAFT_QUALIFIKATIONEN


def erlaubte_dienst_codes(person: dict) -> set[str]:
    feste_dienstart = str(person.get("feste_dienstart", "Alle Dienste")).strip()
    return FESTE_DIENSTART_CODES.get(feste_dienstart, {"F", "S", "N"})


def feste_dienstart_text(person: dict) -> str:
    value = str(person.get("feste_dienstart", "Alle Dienste")).strip()
    return value if value in FESTE_DIENSTARTEN else "Alle Dienste"


def daterange(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def remove_person_from_day(entry: dict, name: str) -> None:
    target = name.casefold()
    for column in CODE_TO_COLUMN.values():
        entry[column] = join_names(
            value
            for value in split_names(entry.get(column, ""))
            if value.casefold() != target
        )


def apply_vorgaben(year: int, month: int, employees, vorgaben):
    plan = create_empty_month(year, month)
    by_day = {parse_date(row["Datum"]): row for row in plan}
    employee_by_id = {int(person["id"]): person for person in employees}
    month_start = date(year, month, 1)
    month_end = date(year, month, calendar.monthrange(year, month)[1])

    # Neuere Vorgaben haben Vorrang, falls versehentlich mehrere Vorgaben kollidieren.
    for vorgabe in sorted(vorgaben or [], key=lambda item: int(item.get("id", 0))):
        person = employee_by_id.get(int(vorgabe.get("mitarbeitende_id", -1)))
        code = str(vorgabe.get("code", "")).strip()
        if not person or code not in VORGABE_LABELS:
            continue
        try:
            start = datetime.strptime(str(vorgabe["startdatum"]), "%Y-%m-%d").date()
            end = datetime.strptime(str(vorgabe["enddatum"]), "%Y-%m-%d").date()
        except (KeyError, TypeError, ValueError):
            continue
        start = max(start, month_start)
        end = min(end, month_end)
        if start > end:
            continue
        for day_value in daterange(start, end):
            entry = by_day.get(day_value)
            if not entry:
                continue
            remove_person_from_day(entry, person["name"])
            column = CODE_TO_COLUMN[code]
            entry[column] = join_names(split_names(entry.get(column, "")) + [person["name"]])
            entry["Status"] = "Abwesenheiten übernommen"
    return plan

def month_from_plan(plan):
    for entry in normalise_plan(plan):
        parsed = parse_date(entry.get("Datum"))
        if parsed:
            return parsed.year, parsed.month
    today = date.today()
    return today.year, today.month


def code_for_person_day(entry, name: str) -> str:
    target = name.casefold()
    mapping = {
        "Frühdienst": "F",
        "Spätdienst": "S",
        "Nachtdienst": "N",
        "Urlaub": "U",
        "Wunschfrei": "W",
        "Fortbildung": "FB",
        "Krank": "K",
        "Frei": "Frei",
    }
    for column, code in mapping.items():
        if any(x.casefold() == target for x in split_names(entry.get(column, ""))):
            return code
    return ""

def workdays_in_month(plan) -> int:
    year, month = month_from_plan(plan)
    return max(
        1,
        sum(
            1
            for day in range(1, calendar.monthrange(year, month)[1] + 1)
            if date(year, month, day).weekday() < 5
        ),
    )


def calculate_hours(plan, employees, settings) -> pd.DataFrame:
    plan = normalise_plan(plan)
    workdays = workdays_in_month(plan)
    rows = []

    for person in employees:
        name = person["name"]
        target = monthly_target(person)
        absence_day_hours = target / workdays if workdays else 0.0
        counts = {"F": 0, "S": 0, "N": 0, "U": 0, "W": 0, "FB": 0, "K": 0, "Frei": 0}
        credited_absence_days = 0

        for entry in plan:
            code = code_for_person_day(entry, name)
            if code in counts:
                counts[code] += 1
            day_value = parse_date(entry.get("Datum"))
            if code in {"U", "FB", "K"} and day_value and day_value.weekday() < 5:
                credited_absence_days += 1

        dienststunden = (
            counts["F"] * float(settings["f_stunden"])
            + counts["S"] * float(settings["s_stunden"])
            + counts["N"] * float(settings["n_stunden"])
        )
        abwesenheitsstunden = credited_absence_days * absence_day_hours
        actual = round(dienststunden + abwesenheitsstunden, 1)

        rows.append(
            {
                "Mitarbeitende": name,
                "Soll": round(target, 1),
                "Ist": actual,
                "+/-": round(actual - target, 1),
                "F": counts["F"],
                "S": counts["S"],
                "N": counts["N"],
                "U": counts["U"],
                "W": counts["W"],
                "FB": counts["FB"],
                "K": counts["K"],
            }
        )

    return pd.DataFrame(rows)

def plan_to_matrix(plan, employees, settings) -> pd.DataFrame:
    """Monatstabelle mit Stunden direkt neben dem Namen.

    Die Reihenfolge ist bewusst: Mitarbeitende, Soll, Ist, Rest und danach
    die einzelnen Kalendertage. So bleibt die Stundenkontrolle immer sichtbar.
    """
    plan = normalise_plan(plan)
    stats = calculate_hours(plan, employees, settings)
    stats_by_name = {
        str(row["Mitarbeitende"]).casefold(): row
        for _, row in stats.iterrows()
    }
    rows = []

    for person in employees:
        name = person["name"]
        stat = stats_by_name.get(name.casefold(), {})
        target = float(stat.get("Soll", monthly_target(person)))
        actual = float(stat.get("Ist", 0.0))
        row = {
            "Mitarbeitende": name,
            "Soll": target,
            "Ist": actual,
            "Rest": round(target - actual, 1),
        }
        for entry in plan:
            parsed = parse_date(entry.get("Datum"))
            if parsed:
                row[day_column(parsed)] = code_for_person_day(entry, name)
        rows.append(row)

    return pd.DataFrame(rows)


def matrix_to_plan(matrix: pd.DataFrame, old_plan):
    old_plan = normalise_plan(old_plan)
    result = []

    code_map = {
        "F": "Frühdienst",
        "S": "Spätdienst",
        "N": "Nachtdienst",
        "NA": "Nachtdienst",
        "U": "Urlaub",
        "W": "Wunschfrei",
        "FB": "Fortbildung",
        "K": "Krank",
        "FREI": "Frei",
    }

    for entry in old_plan:
        parsed = parse_date(entry.get("Datum"))
        if not parsed:
            continue
        column = day_column(parsed)
        assignments = {value: [] for value in CODE_TO_COLUMN.values()}
        for _, row in matrix.iterrows():
            name = str(row.get("Mitarbeitende", "")).strip()
            code = str(row.get(column, "") or "").strip().upper()
            target_column = code_map.get(code)
            if name and target_column:
                assignments[target_column].append(name)

        result.append(
            {
                "Tag": parsed.day,
                "Datum": parsed.strftime("%d.%m.%Y"),
                **{key: join_names(value) for key, value in assignments.items()},
                "Status": "Manuell bearbeitet",
            }
        )
    return result



def replace_name_in_plan(plan, old_name: str, new_name: str):
    target = old_name.casefold()
    updated = normalise_plan(plan)
    for entry in updated:
        for column in CODE_TO_COLUMN.values():
            entry[column] = join_names(
                new_name if value.casefold() == target else value
                for value in split_names(entry.get(column, ""))
            )
    return updated

def remove_name_from_plan(plan, name: str):
    target = name.casefold()
    updated = normalise_plan(plan)
    for entry in updated:
        for column in CODE_TO_COLUMN.values():
            entry[column] = join_names(
                value
                for value in split_names(entry.get(column, ""))
                if value.casefold() != target
            )
    return updated

def shift_interval(day_value: date, code: str, settings: dict):
    prefixes = {"F": "f", "S": "s", "N": "n"}
    prefix = prefixes[code]
    start_dt = datetime.combine(day_value, parse_clock(settings[f"{prefix}_start"]))
    end_dt = datetime.combine(day_value, parse_clock(settings[f"{prefix}_ende"]))
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
    return start_dt, end_dt


def validate_plan(plan, employees, settings):
    errors = []
    warnings = []
    employee_map = {p["name"].casefold(): p for p in employees}
    assignments = {key: [] for key in employee_map}
    work_dates = {key: set() for key in employee_map}
    night_dates = {key: set() for key in employee_map}

    for entry in normalise_plan(plan):
        day_value = parse_date(entry.get("Datum"))
        if not day_value:
            continue

        appearances = {}
        for column, code in [
            ("Frühdienst", "F"), ("Spätdienst", "S"), ("Nachtdienst", "N"),
            ("Urlaub", "U"), ("Wunschfrei", "W"), ("Fortbildung", "FB"),
            ("Krank", "K"), ("Frei", "Frei"),
        ]:
            for name in split_names(entry.get(column, "")):
                key = name.casefold()
                appearances.setdefault(key, []).append(code)
                if key not in employee_map:
                    errors.append(f"Unbekannte Person im Plan: {name}")
                    continue
                if code in {"F", "S", "N"}:
                    person = employee_map[key]
                    if code not in erlaubte_dienst_codes(person):
                        errors.append(
                            f"{name}: {code} passt nicht zur festen Dienstart "
                            f"„{feste_dienstart_text(person)}“ ({day_value.strftime('%d.%m.')})."
                        )
                    if code == "N" and str(person.get("nachtdienst", "Nein")) != "Ja":
                        errors.append(f"{name}: Nachtdienst ist nicht erlaubt ({day_value.strftime('%d.%m.')}).")
                    start_dt, end_dt = shift_interval(day_value, code, settings)
                    assignments[key].append((start_dt, end_dt, code))
                    work_dates[key].add(day_value)
                    if code == "N":
                        night_dates[key].add(day_value)

        for key, codes in appearances.items():
            if len(codes) > 1 and key in employee_map:
                errors.append(
                    f'{employee_map[key]["name"]}: Mehrere Einträge am {day_value.strftime("%d.%m.")}: '
                    + ", ".join(codes)
                )

        actual = {
            "F": len(split_names(entry.get("Frühdienst", ""))),
            "S": len(split_names(entry.get("Spätdienst", ""))),
            "N": len(split_names(entry.get("Nachtdienst", ""))),
        }
        actual_fach = {}
        for code, column in [("F", "Frühdienst"), ("S", "Spätdienst"), ("N", "Nachtdienst")]:
            actual_fach[code] = sum(
                1
                for name in split_names(entry.get(column, ""))
                if name.casefold() in employee_map and is_fachkraft(employee_map[name.casefold()])
            )

        for code, prefix in [("F", "f"), ("S", "s"), ("N", "n")]:
            required = int(settings.get(f"min_{prefix}", 0))
            required_fach = int(settings.get(f"min_fach_{prefix}", 0))
            if actual[code] < required:
                warnings.append(f"{day_value.strftime('%d.%m.')}: {code} nur {actual[code]}/{required} besetzt.")
            if regel_aktiv(settings, "regel_fachkraft_aktiv") and actual_fach[code] < required_fach:
                errors.append(
                    f"{day_value.strftime('%d.%m.')}: {code} nur {actual_fach[code]}/{required_fach} Fachkräfte."
                )

    if regel_aktiv(settings, "regel_ruhezeit_aktiv"):
        min_rest = float(settings.get("ruhezeit", 11.0))
        for key, items in assignments.items():
            items.sort(key=lambda value: value[0])
            name = employee_map[key]["name"]
            for previous, current in zip(items, items[1:]):
                rest = (current[0] - previous[1]).total_seconds() / 3600.0
                if rest < min_rest:
                    errors.append(
                        f"{name}: Nur {rest:.1f} Stunden Ruhezeit zwischen {previous[2]} und {current[2]}."
                    )

    def longest_run(values: set[date]) -> int:
        longest = current = 0
        previous = None
        for value in sorted(values):
            current = current + 1 if previous and value == previous + timedelta(days=1) else 1
            longest = max(longest, current)
            previous = value
        return longest

    max_work = int(settings.get("max_arbeitstage", 6))
    max_nights = int(settings.get("max_nachtdienste", 4))
    if regel_aktiv(settings, "regel_max_arbeitstage_aktiv"):
        for key, dates in work_dates.items():
            if longest_run(dates) > max_work:
                errors.append(f'{employee_map[key]["name"]}: Mehr als {max_work} Arbeitstage hintereinander.')
    if regel_aktiv(settings, "regel_max_nachtdienste_aktiv"):
        for key, dates in night_dates.items():
            if longest_run(dates) > max_nights:
                errors.append(f'{employee_map[key]["name"]}: Mehr als {max_nights} Nachtdienste hintereinander.')

    max_weekends = max(0, int(settings.get("max_wochenenden", 2)))
    no_consecutive_weekends = bool(settings.get("keine_folgewochenenden", True))
    for key, dates in work_dates.items():
        weekends = worked_weekends(dates)
        if regel_aktiv(settings, "regel_max_wochenenden_aktiv") and len(weekends) > max_weekends:
            errors.append(
                f'{employee_map[key]["name"]}: An {len(weekends)} Wochenenden gearbeitet; '
                f'erlaubt sind höchstens {max_weekends}.'
            )
        if no_consecutive_weekends and has_consecutive_weekends(weekends):
            errors.append(
                f'{employee_map[key]["name"]}: Zwei direkt aufeinanderfolgende Wochenenden eingeplant.'
            )

    if regel_aktiv(settings, "regel_sollstunden_aktiv"):
        stats = calculate_hours(plan, employees, settings)
        for _, row in stats.iterrows():
            actual = float(row["Ist"])
            target = float(row["Soll"])
            if actual > target + 0.05:
                errors.append(
                    f'{row["Mitarbeitende"]}: {actual:.1f} Iststunden überschreiten '
                    f'das Soll von {target:.1f} Stunden.'
                )

    employee_by_id = {int(person["id"]): person for person in employees}
    normalised = normalise_plan(plan)
    for rule in eigene_planungsregeln(settings):
        if not bool(rule.get("aktiv", True)):
            continue
        rule_name = str(rule.get("name", "Eigene Regel"))
        rule_type = str(rule.get("typ", ""))

        if rule_type == "nicht_gemeinsam":
            first = employee_by_id.get(regel_zahl(rule, "mitarbeitende_a_id", -1))
            second = employee_by_id.get(regel_zahl(rule, "mitarbeitende_b_id", -1))
            rule_shift = str(rule.get("dienst", "ALL"))
            if not first or not second:
                continue
            for entry in normalised:
                day_value = parse_date(entry.get("Datum"))
                codes = [rule_shift] if rule_shift in {"F", "S", "N"} else ["F", "S", "N"]
                for code in codes:
                    names = {
                        name.casefold()
                        for name in split_names(entry.get(CODE_TO_COLUMN[code], ""))
                    }
                    if first["name"].casefold() in names and second["name"].casefold() in names:
                        errors.append(
                            f'{rule_name}: {first["name"]} und {second["name"]} sind '
                            f'am {day_value.strftime("%d.%m.") if day_value else "unbekannten Tag"} '
                            f'gemeinsam in {code}.'
                        )

        elif rule_type == "frei_nach_dienst":
            trigger_code = str(rule.get("dienst", "N"))
            free_days = max(1, regel_zahl(rule, "freie_tage", 1))
            for person in employees:
                if not regel_gilt_fuer_person(rule, person):
                    continue
                key = person["name"].casefold()
                trigger_days = {
                    item[0].date()
                    for item in assignments[key]
                    if item[2] == trigger_code
                }
                for trigger_day in trigger_days:
                    blocked_days = {
                        trigger_day + timedelta(days=offset)
                        for offset in range(1, free_days + 1)
                    }
                    if blocked_days & work_dates[key]:
                        errors.append(
                            f'{rule_name}: {person["name"]} benötigt nach {trigger_code} '
                            f'am {trigger_day.strftime("%d.%m.")} {free_days} freie Tage.'
                        )

        elif rule_type == "max_dienstfolge":
            target_code = str(rule.get("dienst", "N"))
            maximum = max(1, regel_zahl(rule, "maximum", 1))
            for person in employees:
                if not regel_gilt_fuer_person(rule, person):
                    continue
                key = person["name"].casefold()
                dates = {
                    item[0].date()
                    for item in assignments[key]
                    if item[2] == target_code
                }
                if longest_run(dates) > maximum:
                    errors.append(
                        f'{rule_name}: {person["name"]} hat mehr als {maximum} '
                        f'{target_code}-Dienste hintereinander.'
                    )

    errors = list(dict.fromkeys(errors))
    warnings = list(dict.fromkeys(warnings))
    return errors, warnings

def create_automatic_month(year: int, month: int, employees, settings, vorgaben):
    """Bedarfsorientierte Planung wie in professionellen Personaleinsatzsystemen.

    1. Urlaub, Wunschfrei, Fortbildung und Krank werden zuerst als Abwesenheiten gesperrt.
    2. Der komplette Nachtdienst des Monats wird zuerst geplant.
    3. Danach werden die Wochenenden und erst anschließend die Werktage besetzt.
    4. Pro Person gelten höchstens zwei Arbeitswochenenden; Folgewochenenden sind gesperrt.
    5. Sollstunden und Dienstverteilung dienen zur fairen Auswahl.
    """
    plan = apply_vorgaben(year, month, employees, vorgaben)
    if not employees:
        return plan, {"assigned": 0, "fixed": 0, "shortages": ["Keine Mitarbeitenden vorhanden."]}

    employee_map = {str(person["name"]).casefold(): person for person in employees}
    employee_by_id = {int(person["id"]): person for person in employees}
    employee_order = {str(person["name"]).casefold(): index for index, person in enumerate(employees)}
    planned_hours = {key: 0.0 for key in employee_map}
    assignments = {key: [] for key in employee_map}
    service_counts = {key: {"F": 0, "S": 0, "N": 0} for key in employee_map}
    weekend_sets = {key: set() for key in employee_map}

    # Bezahlte Abwesenheiten zählen bereits zum Monatssoll und müssen vor der
    # Dienstverteilung berücksichtigt werden. Wochenenden erhalten dabei keine
    # pauschalen Urlaubs-/Fortbildungs-/Krankstunden.
    workdays = workdays_in_month(plan)
    for person in employees:
        key = str(person["name"]).casefold()
        absence_day_hours = monthly_target(person) / workdays if workdays else 0.0
        credited_days = 0
        for entry in plan:
            day_value = parse_date(entry.get("Datum"))
            code = code_for_person_day(entry, str(person["name"]))
            if code in {"U", "FB", "K"} and day_value and day_value.weekday() < 5:
                credited_days += 1
        planned_hours[key] = credited_days * absence_day_hours
    shortages = []
    assigned_total = 0
    fixed_total = 0

    code_to_hours = {
        "F": float(settings.get("f_stunden", 0.0)),
        "S": float(settings.get("s_stunden", 0.0)),
        "N": float(settings.get("n_stunden", 0.0)),
    }
    demand = {
        "F": max(0, int(settings.get("min_f", 0))),
        "S": max(0, int(settings.get("min_s", 0))),
        "N": max(0, int(settings.get("min_n", 0))),
    }
    if regel_aktiv(settings, "regel_fachkraft_aktiv"):
        fach_demand = {
            "F": min(demand["F"], max(0, int(settings.get("min_fach_f", 0)))),
            "S": min(demand["S"], max(0, int(settings.get("min_fach_s", 0)))),
            "N": min(demand["N"], max(0, int(settings.get("min_fach_n", 0)))),
        }
    else:
        fach_demand = {"F": 0, "S": 0, "N": 0}
    min_rest = float(settings.get("ruhezeit", 11.0))
    max_work = int(settings.get("max_arbeitstage", 6))
    max_nights = int(settings.get("max_nachtdienste", 4))
    max_weekends = max(0, int(settings.get("max_wochenenden", 2)))
    no_consecutive_weekends = bool(settings.get("keine_folgewochenenden", True))
    active_custom_rules = [
        rule
        for rule in eigene_planungsregeln(settings)
        if bool(rule.get("aktiv", True))
    ]
    days_in_month = len(plan)

    def longest_run(values: set[date]) -> int:
        longest = current = 0
        previous = None
        for value in sorted(values):
            current = current + 1 if previous and value == previous + timedelta(days=1) else 1
            longest = max(longest, current)
            previous = value
        return longest

    def person_code(entry: dict, person: dict) -> str:
        return code_for_person_day(entry, str(person["name"]))

    def violates_custom_rule(person: dict, entry: dict, code: str) -> bool:
        key = str(person["name"]).casefold()
        day_value = parse_date(entry.get("Datum"))
        if day_value is None:
            return True

        for rule in active_custom_rules:
            rule_type = str(rule.get("typ", ""))

            if rule_type == "nicht_gemeinsam":
                first_id = regel_zahl(rule, "mitarbeitende_a_id", -1)
                second_id = regel_zahl(rule, "mitarbeitende_b_id", -1)
                person_id = int(person.get("id", -1))
                if person_id not in {first_id, second_id}:
                    continue
                other_id = second_id if person_id == first_id else first_id
                other = employee_by_id.get(other_id)
                if other is None:
                    continue
                other_code = person_code(entry, other)
                rule_shift = str(rule.get("dienst", "ALL"))
                if other_code == code and (rule_shift == "ALL" or rule_shift == code):
                    return True

            elif rule_type == "frei_nach_dienst" and regel_gilt_fuer_person(rule, person):
                trigger_code = str(rule.get("dienst", "N"))
                free_days = max(1, regel_zahl(rule, "freie_tage", 1))
                for existing_day, _, _, existing_code in assignments[key]:
                    if (
                        existing_code == trigger_code
                        and existing_day < day_value <= existing_day + timedelta(days=free_days)
                    ):
                        return True
                    if (
                        code == trigger_code
                        and day_value < existing_day <= day_value + timedelta(days=free_days)
                    ):
                        return True

            elif rule_type == "max_dienstfolge" and regel_gilt_fuer_person(rule, person):
                target_code = str(rule.get("dienst", "N"))
                maximum = max(1, regel_zahl(rule, "maximum", 1))
                if code == target_code:
                    dates = {
                        item[0]
                        for item in assignments[key]
                        if item[3] == target_code
                    } | {day_value}
                    if longest_run(dates) > maximum:
                        return True

        return False

    def register(person: dict, entry: dict, code: str, fixed: bool = False) -> None:
        nonlocal assigned_total, fixed_total
        key = str(person["name"]).casefold()
        day_value = parse_date(entry["Datum"])
        if not day_value:
            return
        start_dt, end_dt = shift_interval(day_value, code, settings)
        assignments[key].append((day_value, start_dt, end_dt, code))
        assignments[key].sort(key=lambda item: item[1])
        planned_hours[key] += code_to_hours[code]
        service_counts[key][code] += 1
        anchor = weekend_anchor(day_value)
        if anchor is not None:
            weekend_sets[key].add(anchor)
        if fixed:
            fixed_total += 1
        else:
            assigned_total += 1

    # Bereits manuell im geöffneten Plan eingetragene Dienste registrieren.
    for entry in plan:
        for code, column in [("F", "Frühdienst"), ("S", "Spätdienst"), ("N", "Nachtdienst")]:
            for name in split_names(entry.get(column, "")):
                person = employee_map.get(name.casefold())
                if person:
                    register(person, entry, code, fixed=True)

    def can_assign(person: dict, entry: dict, code: str) -> bool:
        key = str(person["name"]).casefold()
        day_value = parse_date(entry["Datum"])
        if day_value is None or monthly_target(person) <= 0 or code_to_hours[code] <= 0:
            return False
        # Sollstunden sind eine feste Obergrenze. Die Automatik lässt einen
        # Dienst offen, statt die Person über ihr persönliches Monatssoll zu planen.
        if (
            regel_aktiv(settings, "regel_sollstunden_aktiv")
            and planned_hours[key] + code_to_hours[code] > monthly_target(person) + 0.001
        ):
            return False
        if person_code(entry, person):
            return False
        # Eine feste Dienstart schränkt nur die erlaubte Schicht ein. Sie erzeugt
        # keine tägliche Einplanung und ändert weder Sollstunden noch freie Tage.
        if code not in erlaubte_dienst_codes(person):
            return False
        if code == "N" and str(person.get("nachtdienst", "Nein")) != "Ja":
            return False

        candidate_start, candidate_end = shift_interval(day_value, code, settings)
        for _, existing_start, existing_end, _ in assignments[key]:
            if candidate_start >= existing_end:
                rest = (candidate_start - existing_end).total_seconds() / 3600.0
            elif existing_start >= candidate_end:
                rest = (existing_start - candidate_end).total_seconds() / 3600.0
            else:
                return False
            if regel_aktiv(settings, "regel_ruhezeit_aktiv") and rest < min_rest:
                return False

        work_days = {item[0] for item in assignments[key]} | {day_value}
        if regel_aktiv(settings, "regel_max_arbeitstage_aktiv") and longest_run(work_days) > max_work:
            return False
        if code == "N" and regel_aktiv(settings, "regel_max_nachtdienste_aktiv"):
            night_days = {item[0] for item in assignments[key] if item[3] == "N"} | {day_value}
            if longest_run(night_days) > max_nights:
                return False

        anchor = weekend_anchor(day_value)
        if anchor is not None:
            candidate_weekends = set(weekend_sets[key]) | {anchor}
            if regel_aktiv(settings, "regel_max_wochenenden_aktiv") and len(candidate_weekends) > max_weekends:
                return False
            if no_consecutive_weekends and has_consecutive_weekends(candidate_weekends):
                return False
        if violates_custom_rule(person, entry, code):
            return False
        return True

    def assign(person: dict, entry: dict, code: str) -> None:
        column = CODE_TO_COLUMN[code]
        entry[column] = join_names(split_names(entry.get(column, "")) + [person["name"]])
        register(person, entry, code, fixed=False)

    night_capable_target = sum(
        monthly_target(person)
        for person in employees
        if str(person.get("nachtdienst", "Nein")) == "Ja" and monthly_target(person) > 0
    )
    total_night_slots = days_in_month * demand["N"]

    only_shift_labels = {
        "F": "Nur Frühdienst",
        "S": "Nur Spätdienst",
        "N": "Nur Nachtdienst",
    }

    def shift_priority(person: dict, code: str) -> int:
        """Mitarbeitende mit nur einer Dienstart werden zuerst eingeplant.

        Dadurch werden zum Beispiel reine Nachtwachen nicht von Mitarbeitenden
        mit „Alle Dienste“ aus den Nachtdiensten verdrängt. Erst wenn diese
        Personen nicht verfügbar sind, nutzt die Automatik flexibel einsetzbare
        Mitarbeitende.
        """
        return 0 if feste_dienstart_text(person) == only_shift_labels[code] else 1

    def weekend_pair_priority(person: dict, day_value: date, code: str) -> int:
        """Am Sonntag möglichst dieselben Personen und Schichten wie Samstag nutzen."""
        if not regel_aktiv(settings, "regel_wochenende_gleich_aktiv"):
            return 0
        anchor = weekend_anchor(day_value)
        if anchor is None:
            return 0

        key = str(person["name"]).casefold()
        same_weekend = [
            item
            for item in assignments[key]
            if weekend_anchor(item[0]) == anchor
        ]
        if any(item[3] == code for item in same_weekend):
            return 0
        if same_weekend:
            return 1
        return 2

    def score_candidate(person: dict, entry: dict, code: str):
        key = str(person["name"]).casefold()
        day_value = parse_date(entry["Datum"])
        target = max(monthly_target(person), 0.01)
        projected = planned_hours[key] + code_to_hours[code]
        projected_utilisation = projected / target
        over_ratio = max(0.0, projected - target) / target
        previous_night = any(
            item[0] == day_value - timedelta(days=1) and item[3] == "N"
            for item in assignments[key]
        )
        next_night = any(
            item[0] == day_value + timedelta(days=1) and item[3] == "N"
            for item in assignments[key]
        )
        night_block_penalty = (
            0
            if (
                not regel_aktiv(settings, "regel_nachtblock_aktiv")
                or code != "N"
                or previous_night
                or next_night
            )
            else 1
        )
        expected_nights = (
            total_night_slots * target / night_capable_target
            if code == "N" and night_capable_target > 0
            else 1.0
        )
        night_load = (
            (service_counts[key]["N"] + 1) / max(expected_nights, 0.5)
            if code == "N"
            else 0.0
        )
        weekend_load = len(weekend_sets[key])
        weekend_pair = weekend_pair_priority(person, day_value, code)
        restricted_shift = shift_priority(person, code)
        rotation = (day_value.day - 1 - employee_order[key] * 2) % max(1, days_in_month)

        # Für Nachtdienste gilt eine feste Priorität:
        # 1. Mitarbeitende mit „Nur Nachtdienst"
        # 2. bereits begonnene Nachtblöcke fortsetzen
        # 3. innerhalb dieser Gruppen fair nach Sollstunden verteilen
        if code == "N":
            return (
                restricted_shift,
                night_block_penalty,
                round(projected_utilisation, 6),
                round(night_load, 6),
                round(over_ratio, 6),
                service_counts[key][code],
                sum(service_counts[key].values()),
                rotation,
                str(person["name"]).casefold(),
            )

        # Früh- und Spätdienste werden erst nach allen Nachtdiensten geplant.
        # Am Wochenende bleiben die Personen nach Möglichkeit an beiden Tagen
        # in derselben Schicht.
        return (
            restricted_shift,
            weekend_pair,
            round(projected_utilisation, 6),
            round(over_ratio, 6),
            weekend_load if day_value.weekday() >= 5 else 0,
            service_counts[key][code],
            sum(service_counts[key].values()),
            rotation,
            str(person["name"]).casefold(),
        )

    def fill_slots(entry: dict, code: str, needed: int, fach_only: bool) -> None:
        for _ in range(max(0, needed)):
            candidates = [
                person for person in employees
                if (not fach_only or is_fachkraft(person)) and can_assign(person, entry, code)
            ]
            if not candidates:
                kind = "Fachkraft" if fach_only else "Person"
                shortages.append(f'{entry["Datum"]}: {code} – keine passende {kind} verfügbar.')
                return
            selected = min(candidates, key=lambda person: score_candidate(person, entry, code))
            assign(selected, entry, code)

    def fill_fach(entry: dict, code: str) -> None:
        current_names = split_names(entry.get(CODE_TO_COLUMN[code], ""))
        current_fach = sum(
            1 for name in current_names
            if name.casefold() in employee_map and is_fachkraft(employee_map[name.casefold()])
        )
        fill_slots(entry, code, fach_demand[code] - current_fach, fach_only=True)

    def fill_total(entry: dict, code: str) -> None:
        current_total = len(split_names(entry.get(CODE_TO_COLUMN[code], "")))
        fill_slots(entry, code, demand[code] - current_total, fach_only=False)

    weekend_entries = [
        entry for entry in plan
        if (day_value := parse_date(entry.get("Datum"))) and day_value.weekday() >= 5
    ]
    weekday_entries = [
        entry for entry in plan
        if (day_value := parse_date(entry.get("Datum"))) and day_value.weekday() < 5
    ]
    planning_days = (
        weekend_entries + weekday_entries
        if regel_aktiv(settings, "regel_wochenende_zuerst_aktiv")
        else list(plan)
    )

    if regel_aktiv(settings, "regel_nacht_zuerst_aktiv"):
        for entry in plan:
            fill_fach(entry, "N")
            fill_total(entry, "N")
        for entry in planning_days:
            for code in ("F", "S"):
                fill_fach(entry, code)
                fill_total(entry, code)
            entry["Status"] = "Nachtdienste zuerst geplant"
    else:
        for entry in planning_days:
            for code in ("F", "S", "N"):
                fill_fach(entry, code)
                fill_total(entry, code)
            entry["Status"] = "Nach aktiven Planungsregeln geplant"

    required_hours = days_in_month * sum(demand[c] * code_to_hours[c] for c in ("F", "S", "N"))
    available_hours = sum(monthly_target(person) for person in employees)
    return plan, {
        "assigned": assigned_total,
        "fixed": fixed_total,
        "shortages": list(dict.fromkeys(shortages)),
        "required_hours": round(required_hours, 1),
        "available_hours": round(available_hours, 1),
    }

def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9ÄÖÜäöüß_-]+", "_", str(name).strip())
    return cleaned.strip("_") or "Dienstplan"


def export_matrix(plan, employees, settings):
    return plan_to_matrix(plan, employees, settings)


def create_excel(plan, plan_name, employees, settings):
    matrix = export_matrix(plan, employees, settings)
    year, month = month_from_plan(plan)
    output = BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Dienstplan"

    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(matrix.columns))
    title = sheet.cell(1, 1, plan_name)
    title.font = Font(size=16, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="203040")
    title.alignment = Alignment(horizontal="center")

    for column_index, column in enumerate(matrix.columns, start=1):
        cell = sheet.cell(3, column_index, column)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="44546A")
        cell.alignment = Alignment(horizontal="center")

    for row_index, row in enumerate(matrix.itertuples(index=False), start=4):
        for column_index, value in enumerate(row, start=1):
            cell = sheet.cell(row_index, column_index, value)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for column_index, column in enumerate(matrix.columns, start=1):
        width = 22 if column == "Mitarbeitende" else 7
        if re.match(r"^\d{2} ", str(column)):
            day_number = int(str(column)[:2])
            if date(year, month, day_number).weekday() >= 5:
                for row_index in range(3, len(matrix) + 4):
                    sheet.cell(row_index, column_index).fill = PatternFill("solid", fgColor="FFF2CC")
        sheet.column_dimensions[get_column_letter(column_index)].width = width

    workbook.save(output)
    output.seek(0)
    return output.getvalue()


def create_pdf(plan, plan_name, employees, settings):
    matrix = export_matrix(plan, employees, settings)
    output = BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A3),
        leftMargin=7 * mm,
        rightMargin=7 * mm,
        topMargin=7 * mm,
        bottomMargin=7 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#172033"),
    )
    cell_style = ParagraphStyle(
        "Cell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=5.7,
        leading=6.4,
        alignment=1,
    )
    header_style = ParagraphStyle(
        "Header",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )

    data = [[Paragraph(escape(str(column)), header_style) for column in matrix.columns]]
    for _, row in matrix.iterrows():
        data.append([Paragraph(escape(str(value)), cell_style) for value in row])

    widths = [32 * mm if column == "Mitarbeitende" else 8.2 * mm for column in matrix.columns]
    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#44546A")),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#C8CED8")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    document.build([Paragraph(escape(plan_name), title_style), table])
    output.seek(0)
    return output.getvalue()


# -----------------------------------------------------------------------------
# Neue Benutzeroberfläche
# -----------------------------------------------------------------------------

NAVIGATION = {
    "🏠  Übersicht": "Übersicht",
    "👥  Mitarbeitende": "Mitarbeitende",
    "📅  Abwesenheiten": "Abwesenheiten",
    "📐  Planungsregeln": "Planungsregeln",
    "🧩  Vorlagen": "Vorlagen",
    "🗓️  Dienstplan": "Dienstplan",
    "📊  Auswertung": "Auswertung",
    "💬  Kontakt": "Feedback & Hilfe",
    "⚙️  Einstellungen": "Einstellungen",
}

BUILTIN_TEMPLATES = {
    "Standardstation": {
        "description": "Ausgewogene Grundbesetzung für eine normale Pflegestation.",
        "settings": {
            "min_f": 4,
            "min_s": 3,
            "min_n": 1,
            "min_fach_f": 1,
            "min_fach_s": 1,
            "min_fach_n": 1,
            "max_arbeitstage": 6,
            "max_nachtdienste": 3,
            "max_wochenenden": 2,
            "keine_folgewochenenden": True,
        },
    },
    "Hoher Pflegebedarf": {
        "description": "Mehr Personal im Früh- und Spätdienst sowie mindestens zwei Fachkräfte am Tag.",
        "settings": {
            "min_f": 5,
            "min_s": 3,
            "min_n": 1,
            "min_fach_f": 2,
            "min_fach_s": 2,
            "min_fach_n": 1,
            "max_arbeitstage": 6,
            "max_nachtdienste": 3,
            "max_wochenenden": 2,
            "keine_folgewochenenden": True,
        },
    },
    "Kleine Station": {
        "description": "Reduzierte Mindestbesetzung für kleinere Bereiche oder Übergangszeiten.",
        "settings": {
            "min_f": 3,
            "min_s": 2,
            "min_n": 1,
            "min_fach_f": 1,
            "min_fach_s": 1,
            "min_fach_n": 1,
            "max_arbeitstage": 6,
            "max_nachtdienste": 3,
            "max_wochenenden": 2,
            "keine_folgewochenenden": True,
        },
    },
}


def section_title(title: str, help_text: str = "") -> None:
    st.markdown(f'<div class="np-section-title">{escape(title)}</div>', unsafe_allow_html=True)
    if help_text:
        st.markdown(f'<div class="np-help">{escape(help_text)}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="np-kpi">
            <div class="np-kpi-label">{escape(label)}</div>
            <div class="np-kpi-value">{escape(value)}</div>
            <div class="np-kpi-note">{escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def german_date(value: str) -> str:
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return str(value)


def german_datetime(value: str) -> str:
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return parsed.strftime("%d.%m.%Y · %H:%M Uhr")
    except (TypeError, ValueError):
        return str(value)


def render_employee_table(items: list[dict]) -> None:
    if not items:
        st.info("Noch keine Mitarbeitenden angelegt.")
        return

    rows = []
    for person in items:
        night_class = "np-badge-green" if str(person.get("nachtdienst")) == "Ja" else "np-badge-gray"
        night_text = "Erlaubt" if str(person.get("nachtdienst")) == "Ja" else "Nicht erlaubt"
        fixed_shift = feste_dienstart_text(person)
        shift_class = "np-badge-blue" if fixed_shift == "Alle Dienste" else "np-badge-orange"
        rows.append(
            "<tr>"
            f'<td><span class="np-name">{escape(str(person["name"]))}</span></td>'
            f'<td>{escape(str(person.get("qualifikation", "")))}</td>'
            f'<td>{escape(percentage_text(monthly_target(person)))}</td>'
            f'<td>{monthly_target(person):.1f} Std.</td>'
            f'<td><span class="np-badge {night_class}">{night_text}</span></td>'
            f'<td><span class="np-badge {shift_class}">{escape(fixed_shift)}</span></td>'
            "</tr>"
        )

    st.markdown(
        """
        <div class="np-table-wrap">
            <table class="np-table">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Qualifikation</th>
                        <th>Beschäftigung</th>
                        <th>Sollstunden</th>
                        <th>Nachtdienst</th>
                        <th>Erlaubte Dienste</th>
                    </tr>
                </thead>
                <tbody>
        """
        + "".join(rows)
        + """
                </tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def readable_plan_matrix(matrix: pd.DataFrame):
    day_columns = [column for column in matrix.columns if re.match(r"^\d{2} ", str(column))]

    def style_code(value):
        code = str(value or "").strip().upper()
        styles = {
            "F": "background-color: #e8f3ff; color: #174ea6; font-weight: 700; text-align: center;",
            "S": "background-color: #fff2df; color: #9a4f00; font-weight: 700; text-align: center;",
            "N": "background-color: #ede9fe; color: #5b21b6; font-weight: 700; text-align: center;",
            "U": "background-color: #e8f8f2; color: #087a54; font-weight: 700; text-align: center;",
            "W": "background-color: #f1f5f9; color: #475467; font-weight: 700; text-align: center;",
            "FB": "background-color: #e0f2fe; color: #075985; font-weight: 700; text-align: center;",
            "K": "background-color: #feecec; color: #b42318; font-weight: 700; text-align: center;",
            "FREI": "background-color: #f4f4f5; color: #52525b; font-weight: 650; text-align: center;",
        }
        return styles.get(code, "text-align: center;")

    styled = matrix.style
    if day_columns:
        styled = styled.map(style_code, subset=day_columns)
    numeric = [column for column in ["Soll", "Ist", "Rest"] if column in matrix.columns]
    if numeric:
        styled = styled.format({column: "{:.1f}" for column in numeric})
    return styled


def settings_snapshot(values: dict) -> dict:
    keys = [
        "vollzeit_monatssoll",
        "f_start", "f_ende", "f_stunden", "min_f", "min_fach_f",
        "s_start", "s_ende", "s_stunden", "min_s", "min_fach_s",
        "n_start", "n_ende", "n_stunden", "min_n", "min_fach_n",
        "ruhezeit", "max_arbeitstage", "max_nachtdienste", "max_wochenenden",
        "keine_folgewochenenden", "max_plusstunden",
        "regel_ruhezeit_aktiv", "regel_max_arbeitstage_aktiv",
        "regel_max_nachtdienste_aktiv", "regel_max_wochenenden_aktiv",
        "regel_sollstunden_aktiv", "regel_fachkraft_aktiv",
        "regel_nacht_zuerst_aktiv", "regel_nachtblock_aktiv",
        "regel_wochenende_zuerst_aktiv", "regel_wochenende_gleich_aktiv",
        "eigene_planungsregeln",
    ]
    return {key: values.get(key) for key in keys if key in values}


def apply_template_settings(template_values: dict) -> None:
    new_settings = {**st.session_state.settings, **template_values}
    save_einstellungen(new_settings)
    st.session_state.settings = new_settings


def copy_plan_to_month(source_plan, year: int, month: int, employees, vorgaben):
    """Kopiert Dienste nach Kalendertag. Neue Abwesenheiten haben Vorrang."""
    target_plan = apply_vorgaben(year, month, employees, vorgaben)
    employee_names = {str(person["name"]).casefold(): str(person["name"]) for person in employees}
    source_by_day = {int(row.get("Tag", index)): row for index, row in enumerate(normalise_plan(source_plan), start=1)}

    for target_entry in target_plan:
        source_entry = source_by_day.get(int(target_entry["Tag"]))
        if not source_entry:
            continue
        for column in ["Frühdienst", "Spätdienst", "Nachtdienst", "Frei"]:
            for source_name in split_names(source_entry.get(column, "")):
                real_name = employee_names.get(source_name.casefold())
                if not real_name:
                    continue
                if code_for_person_day(target_entry, real_name):
                    continue
                target_entry[column] = join_names(split_names(target_entry.get(column, "")) + [real_name])
        target_entry["Status"] = "Aus Vorlage übernommen"
    return target_plan


def show_plan_legend() -> None:
    st.markdown(
        """
        <div style="display:flex;gap:.4rem;flex-wrap:wrap;margin:.25rem 0 1rem 0;">
            <span class="np-badge np-badge-blue">F · Frühdienst</span>
            <span class="np-badge np-badge-orange">S · Spätdienst</span>
            <span class="np-badge" style="background:#ede9fe;color:#5b21b6;">N · Nachtdienst</span>
            <span class="np-badge np-badge-green">U · Urlaub</span>
            <span class="np-badge np-badge-gray">W · Wunschfrei</span>
            <span class="np-badge" style="background:#e0f2fe;color:#075985;">FB · Fortbildung</span>
            <span class="np-badge np-badge-red">K · Krank</span>
            <span class="np-badge np-badge-gray">Frei · Arbeitsfrei</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.sidebar.markdown(
    """
    <div class="np-brand">
        <div class="np-brand-name">🏥 NursePlan Pro</div>
        <div class="np-brand-subtitle">Dienstplanung für Pflegeteams</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.caption(f"Angemeldet: {current_user_email()}")
if st.sidebar.button("Abmelden", width="stretch"):
    st.session_state.auth_logout_requested = True
    delete_auth_cookie()
    try:
        sign_out_user()
    finally:
        for state_key in list(st.session_state):
            if state_key != "auth_logout_requested":
                del st.session_state[state_key]
    st.rerun()

selected_navigation = st.sidebar.radio(
    "Navigation",
    list(NAVIGATION),
    label_visibility="collapsed",
    key="main_navigation",
)
page = NAVIGATION[selected_navigation]
st.sidebar.divider()
st.sidebar.caption("Tipp: Zuerst Mitarbeitende und Abwesenheiten pflegen. Danach den Dienstplan erstellen.")

employees = get_mitarbeitende()
vorgaben = get_vorgaben()
settings = st.session_state.settings
saved_plans = get_dienstplaene()
saved_templates = get_vorlagen()


if page == "Übersicht":
    page_header(
        "Übersicht",
        "Alle wichtigen Informationen auf einen Blick. Von hier aus erkennst du, was vor der Planung noch fehlt.",
    )

    open_plan = st.session_state.dienstplan_vorlage
    open_plan_name = st.session_state.dienstplan_name if open_plan else "Kein Plan geöffnet"
    active_absences = len(vorgaben)
    night_capable = sum(1 for person in employees if str(person.get("nachtdienst")) == "Ja")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Mitarbeitende", str(len(employees)), "im Team gespeichert")
    with c2:
        kpi_card("Nachtdienst möglich", str(night_capable), "Mitarbeitende")
    with c3:
        kpi_card("Abwesenheiten", str(active_absences), "gespeicherte Zeiträume")
    with c4:
        kpi_card("Gespeicherte Pläne", str(len(saved_plans)), "jederzeit wieder ladbar")

    left, right = st.columns([1.35, 1], gap="large")
    with left:
        section_title("Aktueller Arbeitsstand", "Hier siehst du, ob bereits ein Monatsplan geöffnet ist.")
        if open_plan:
            plan_year, plan_month = month_from_plan(open_plan)
            errors, warnings = validate_plan(open_plan, employees, settings)
            st.markdown(
                f"""
                <div class="np-card">
                    <div style="font-size:.8rem;color:#667085;font-weight:700;">GEÖFFNETER PLAN</div>
                    <div style="font-size:1.25rem;font-weight:800;margin:.2rem 0;">{escape(open_plan_name)}</div>
                    <div style="color:#667085;">{MONATSNAMEN[plan_month]} {plan_year} · {len(errors)} Fehler · {len(warnings)} Hinweise</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("Noch kein Dienstplan geöffnet. Öffne die Seite „Dienstplan“, um einen Monat zu erstellen.")

        section_title("Vor der automatischen Planung prüfen")
        checklist = [
            (bool(employees), "Mitarbeitende sind angelegt"),
            (night_capable > 0, "Mindestens eine Person darf Nachtdienst übernehmen"),
            (settings.get("min_f", 0) > 0, "Mindestbesetzung ist eingestellt"),
            (True, "Urlaub und Wünsche wurden geprüft"),
        ]
        for completed, text in checklist:
            icon = "✅" if completed else "⚠️"
            st.write(f"{icon} {text}")

    with right:
        section_title("Aktuelle Besetzung", "Diese Werte verwendet die automatische Planung.")
        st.markdown(
            f"""
            <div class="np-card">
                <div style="display:grid;grid-template-columns:1fr auto;gap:.7rem;">
                    <span>Frühdienst</span><strong>{int(settings.get('min_f', 0))} Personen</strong>
                    <span>Spätdienst</span><strong>{int(settings.get('min_s', 0))} Personen</strong>
                    <span>Nachtdienst</span><strong>{int(settings.get('min_n', 0))} Person(en)</strong>
                    <span>Max. Arbeitstage</span><strong>{int(settings.get('max_arbeitstage', 0))}</strong>
                    <span>Max. Wochenenden</span><strong>{int(settings.get('max_wochenenden', 0))}</strong>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        section_title("Nächste Schritte")
        st.markdown(
            """
            <div class="np-card">
                <strong>1.</strong> Teamdaten prüfen<br>
                <strong>2.</strong> Urlaub und Frei-Wünsche eintragen<br>
                <strong>3.</strong> Vorlage oder Besetzung auswählen<br>
                <strong>4.</strong> Dienstplan erstellen und prüfen
            </div>
            """,
            unsafe_allow_html=True,
        )

    if employees:
        section_title("Team nach Qualifikation")
        qualification_counts = (
            pd.DataFrame(employees)["qualifikation"]
            .fillna("Nicht angegeben")
            .value_counts()
            .rename_axis("Qualifikation")
            .reset_index(name="Anzahl")
        )
        st.bar_chart(qualification_counts.set_index("Qualifikation"))


elif page == "Mitarbeitende":
    page_header(
        "Mitarbeitende",
        "Lege das Team einmal sauber an. Sollstunden, Qualifikation und erlaubte Dienste werden später automatisch berücksichtigt.",
    )

    total_target = sum(monthly_target(person) for person in employees)
    specialists = sum(1 for person in employees if is_fachkraft(person))
    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Teamgröße", str(len(employees)), "aktive Mitarbeitende")
    with c2:
        kpi_card("Fachkräfte", str(specialists), "für die Mindestbesetzung")
    with c3:
        kpi_card("Monatssoll gesamt", f"{total_target:.1f}", "Stunden im Team")

    add_tab, edit_tab, overview_tab = st.tabs(["➕ Neue Person", "✏️ Bearbeiten", "👥 Teamübersicht"])

    with add_tab:
        section_title("Neue Person anlegen", "Trage nur die Angaben ein, die für die Dienstplanung wirklich benötigt werden.")
        with st.form("new_employee_form", clear_on_submit=True):
            left, right = st.columns(2, gap="large")
            with left:
                new_name = st.text_input("Name", placeholder="Zum Beispiel: Anna Müller")
                new_target = st.number_input(
                    "Sollstunden im Monat",
                    min_value=0.0,
                    max_value=300.0,
                    value=float(settings.get("vollzeit_monatssoll", 167.0)),
                    step=0.5,
                    help="Die Stunden laut Arbeitsvertrag für den ausgewählten Monat.",
                )
                st.markdown(
                    f"""
                    <div class="np-percent">
                        <div class="np-percent-label">Berechneter Beschäftigungsumfang</div>
                        <div class="np-percent-value">{percentage_text(new_target)}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with right:
                new_qualification = st.selectbox(
                    "Qualifikation",
                    QUALIFIKATIONEN,
                    help="Die Qualifikation wird für die Fachkraftbesetzung geprüft.",
                )
                new_fixed_shift = st.selectbox(
                    "Welche Dienste sind erlaubt?",
                    FESTE_DIENSTARTEN,
                    help="Wähle „Alle Dienste“, wenn Früh-, Spät- und Nachtdienst möglich sind.",
                )
                new_night = st.checkbox(
                    "Nachtdienst ist erlaubt",
                    value=True,
                    help="Deaktivieren, wenn die Person grundsätzlich keine Nachtdienste übernehmen darf.",
                )
            submitted = st.form_submit_button("Mitarbeitende speichern", type="primary", width="stretch")

        if submitted:
            if not new_name.strip():
                st.warning("Bitte trage einen Namen ein.")
            elif new_target <= 0:
                st.warning("Die Sollstunden müssen größer als 0 sein.")
            elif new_fixed_shift == "Nur Nachtdienst" and not new_night:
                st.warning("Bei „Nur Nachtdienst“ muss Nachtdienst erlaubt sein.")
            elif any(person["name"].strip().casefold() == new_name.strip().casefold() for person in employees):
                st.warning("Dieser Name ist bereits vorhanden.")
            else:
                add_mitarbeitende(
                    name=new_name.strip(),
                    stellenanteil=stored_percentage(new_target),
                    wochenstunden=weekly_from_month(new_target),
                    qualifikation=new_qualification,
                    erfahrung="Nicht erfasst",
                    nachtdienst="Ja" if new_night else "Nein",
                    monatssollstunden=float(new_target),
                    schutzstatus="Keiner",
                    feste_dienstart=new_fixed_shift,
                )
                st.success(f"{new_name.strip()} wurde gespeichert.")
                st.rerun()

    with edit_tab:
        if not employees:
            st.info("Lege zuerst eine Person an.")
        else:
            labels = {person["name"]: person for person in employees}
            selected_name = st.selectbox("Mitarbeitende auswählen", list(labels), key="employee_to_edit")
            selected = labels[selected_name]

            with st.form(f'edit_employee_form_{selected["id"]}'):
                left, right = st.columns(2, gap="large")
                with left:
                    edit_name = st.text_input("Name", value=selected["name"])
                    edit_target = st.number_input(
                        "Sollstunden im Monat",
                        min_value=0.0,
                        max_value=300.0,
                        value=float(monthly_target(selected)),
                        step=0.5,
                    )
                    st.markdown(
                        f"""
                        <div class="np-percent">
                            <div class="np-percent-label">Berechneter Beschäftigungsumfang</div>
                            <div class="np-percent-value">{percentage_text(edit_target)}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with right:
                    qualification_index = (
                        QUALIFIKATIONEN.index(selected["qualifikation"])
                        if selected.get("qualifikation") in QUALIFIKATIONEN
                        else 0
                    )
                    edit_qualification = st.selectbox("Qualifikation", QUALIFIKATIONEN, index=qualification_index)
                    current_fixed_shift = feste_dienstart_text(selected)
                    edit_fixed_shift = st.selectbox(
                        "Welche Dienste sind erlaubt?",
                        FESTE_DIENSTARTEN,
                        index=FESTE_DIENSTARTEN.index(current_fixed_shift),
                    )
                    edit_night = st.checkbox(
                        "Nachtdienst ist erlaubt",
                        value=str(selected.get("nachtdienst", "Nein")) == "Ja",
                    )
                save_employee = st.form_submit_button("Änderungen speichern", type="primary", width="stretch")

            if save_employee:
                if not edit_name.strip():
                    st.warning("Der Name darf nicht leer sein.")
                elif edit_target <= 0:
                    st.warning("Die Sollstunden müssen größer als 0 sein.")
                elif edit_fixed_shift == "Nur Nachtdienst" and not edit_night:
                    st.warning("Bei „Nur Nachtdienst“ muss Nachtdienst erlaubt sein.")
                elif any(
                    person["id"] != selected["id"]
                    and person["name"].strip().casefold() == edit_name.strip().casefold()
                    for person in employees
                ):
                    st.warning("Dieser Name ist bereits vorhanden.")
                else:
                    update_mitarbeitende(
                        mitarbeitende_id=int(selected["id"]),
                        name=edit_name.strip(),
                        stellenanteil=stored_percentage(edit_target),
                        wochenstunden=weekly_from_month(edit_target),
                        qualifikation=edit_qualification,
                        erfahrung=str(selected.get("erfahrung", "Nicht erfasst")),
                        nachtdienst="Ja" if edit_night else "Nein",
                        monatssollstunden=float(edit_target),
                        schutzstatus=str(selected.get("schutzstatus", "Keiner")),
                        feste_dienstart=edit_fixed_shift,
                    )
                    if selected["name"] != edit_name.strip():
                        st.session_state.dienstplan_vorlage = replace_name_in_plan(
                            st.session_state.dienstplan_vorlage,
                            selected["name"],
                            edit_name.strip(),
                        )
                    st.success("Die Änderungen wurden gespeichert.")
                    st.rerun()

            st.divider()
            st.markdown("**Person löschen**")
            confirm_delete = st.checkbox(
                f'Ich möchte {selected["name"]} wirklich löschen.',
                key=f'confirm_delete_{selected["id"]}',
            )
            if st.button("Person löschen", disabled=not confirm_delete, width="stretch"):
                delete_mitarbeitende(int(selected["id"]))
                st.session_state.dienstplan_vorlage = remove_name_from_plan(
                    st.session_state.dienstplan_vorlage,
                    selected["name"],
                )
                st.rerun()

    with overview_tab:
        filter_left, filter_right = st.columns([1.3, 1])
        with filter_left:
            search_name = st.text_input("Suchen", placeholder="Name eingeben", key="employee_search")
        with filter_right:
            qualification_filter = st.selectbox(
                "Qualifikation filtern",
                ["Alle"] + QUALIFIKATIONEN,
                key="employee_qualification_filter",
            )
        filtered = [
            person for person in employees
            if (not search_name.strip() or search_name.strip().casefold() in person["name"].casefold())
            and (qualification_filter == "Alle" or person.get("qualifikation") == qualification_filter)
        ]
        render_employee_table(filtered)
        st.caption(f"Angezeigt: {len(filtered)} von {len(employees)} Mitarbeitenden")


elif page == "Abwesenheiten":
    page_header(
        "Abwesenheiten und Wünsche",
        "Trage Urlaub, Wunschfrei, Fortbildung oder Krankheit ein. Diese Zeiten werden bei der automatischen Planung geschützt.",
    )

    if not employees:
        st.warning("Bitte lege zuerst Mitarbeitende an.")
    else:
        add_tab, overview_tab = st.tabs(["➕ Eintragen", "📋 Übersicht"])
        with add_tab:
            section_title("Neue Abwesenheit", "Wähle eine Person, die Art und den genauen Zeitraum.")
            with st.form("absence_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                employee_options = {person["name"]: int(person["id"]) for person in employees}
                label_to_code = {label: code for code, label in VORGABE_LABELS.items()}
                with c1:
                    selected_employee = st.selectbox("Mitarbeitende", list(employee_options))
                    selected_label = st.selectbox("Art der Abwesenheit", list(label_to_code))
                with c2:
                    start_value = st.date_input("Von", value=date.today(), key="absence_start")
                    end_value = st.date_input("Bis einschließlich", value=date.today(), key="absence_end")
                note = st.text_input("Bemerkung", placeholder="Optional, zum Beispiel: genehmigt")
                save_absence = st.form_submit_button("Abwesenheit speichern", type="primary", width="stretch")

            if save_absence:
                if end_value < start_value:
                    st.warning("Das Enddatum darf nicht vor dem Startdatum liegen.")
                else:
                    add_vorgabe(
                        mitarbeitende_id=employee_options[selected_employee],
                        code=label_to_code[selected_label],
                        startdatum=start_value.isoformat(),
                        enddatum=end_value.isoformat(),
                        bemerkung=note.strip(),
                    )
                    st.success("Die Abwesenheit wurde gespeichert.")
                    st.rerun()

        with overview_tab:
            if not vorgaben:
                st.info("Noch keine Abwesenheiten gespeichert.")
            else:
                employee_names = {int(person["id"]): person["name"] for person in employees}
                overview_rows = []
                delete_options = {}
                for item in vorgaben:
                    person_name = employee_names.get(int(item["mitarbeitende_id"]), "Unbekannt")
                    label = VORGABE_LABELS.get(
                        str(item["code"]),
                        LEGACY_VORGABE_LABELS.get(str(item["code"]), str(item["code"])),
                    )
                    overview_rows.append(
                        {
                            "Mitarbeitende": person_name,
                            "Art": label,
                            "Von": german_date(item["startdatum"]),
                            "Bis": german_date(item["enddatum"]),
                            "Bemerkung": item.get("bemerkung", ""),
                        }
                    )
                    delete_options[
                        f'{person_name} · {label} · {german_date(item["startdatum"])} bis {german_date(item["enddatum"])}'
                    ] = int(item["id"])

                st.dataframe(
                    pd.DataFrame(overview_rows),
                    hide_index=True,
                    width="stretch",
                    column_config={
                        "Mitarbeitende": st.column_config.TextColumn("Mitarbeitende", width="medium"),
                        "Art": st.column_config.TextColumn("Art", width="small"),
                        "Von": st.column_config.TextColumn("Von", width="small"),
                        "Bis": st.column_config.TextColumn("Bis", width="small"),
                        "Bemerkung": st.column_config.TextColumn("Bemerkung", width="large"),
                    },
                )
                st.divider()
                selected_absence = st.selectbox("Eintrag auswählen", list(delete_options))
                if st.button("Ausgewählten Eintrag löschen", width="stretch"):
                    delete_vorgabe(delete_options[selected_absence])
                    st.rerun()


elif page == "Vorlagen":
    page_header(
        "Vorlagen",
        "Mit Vorlagen übernimmst du Besetzung und Planungsregeln. Eigene Einstellungen kannst du jederzeit speichern und wiederverwenden.",
    )

    section_title("Schnellvorlagen", "Eine Schnellvorlage ändert nur Besetzung und Planungsregeln. Mitarbeitende und Abwesenheiten bleiben erhalten.")
    template_columns = st.columns(3)
    for column, (template_name, template_data) in zip(template_columns, BUILTIN_TEMPLATES.items()):
        with column:
            values = template_data["settings"]
            st.markdown(
                f"""
                <div class="np-template-card">
                    <div class="np-template-title">{escape(template_name)}</div>
                    <div class="np-template-text">{escape(template_data['description'])}</div>
                    <hr style="margin:.75rem 0;">
                    <div class="np-template-text">
                        Früh {values['min_f']} · Spät {values['min_s']} · Nacht {values['min_n']}<br>
                        Max. {values['max_arbeitstage']} Arbeitstage · {values['max_wochenenden']} Wochenenden
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(f'„{template_name}“ verwenden', key=f'builtin_{template_name}', width="stretch"):
                apply_template_settings(values)
                st.success(f"Die Vorlage „{template_name}“ wurde übernommen.")
                st.rerun()

    st.divider()
    own_tab, saved_tab = st.tabs(["💾 Eigene Vorlage speichern", "📚 Gespeicherte Vorlagen"])

    with own_tab:
        section_title("Aktuelle Einstellungen als Vorlage speichern")
        with st.form("save_template_form"):
            template_name = st.text_input("Name der Vorlage", placeholder="Zum Beispiel: Station C4 Standard")
            template_description = st.text_area(
                "Kurze Beschreibung",
                placeholder="Wann soll diese Vorlage verwendet werden?",
            )
            include_open_plan = st.checkbox(
                "Geöffneten Monatsplan ebenfalls speichern",
                value=False,
                disabled=not bool(st.session_state.dienstplan_vorlage),
                help="Damit kannst du später auch die Dienste nach Kalendertag in einen neuen Monat übernehmen.",
            )
            save_custom_template = st.form_submit_button("Vorlage speichern", type="primary", width="stretch")

        if save_custom_template:
            if not template_name.strip():
                st.warning("Bitte gib der Vorlage einen Namen.")
            else:
                payload = {
                    "settings": settings_snapshot(settings),
                    "plan": normalise_plan(st.session_state.dienstplan_vorlage) if include_open_plan else [],
                }
                save_vorlage(
                    name=template_name.strip(),
                    beschreibung=template_description.strip(),
                    kategorie="Monatsplan" if include_open_plan else "Besetzung & Regeln",
                    daten=payload,
                )
                st.success("Die Vorlage wurde gespeichert.")
                st.rerun()

    with saved_tab:
        if not saved_templates:
            st.info("Noch keine eigenen Vorlagen gespeichert.")
        else:
            for template in saved_templates:
                with st.expander(f'{template["name"]} · {template["kategorie"]}'):
                    st.write(template.get("beschreibung") or "Keine Beschreibung hinterlegt.")
                    loaded_template = load_vorlage(int(template["id"]))
                    template_payload = loaded_template.get("daten", {}) if loaded_template else {}
                    settings_values = template_payload.get("settings", {})
                    plan_values = template_payload.get("plan", [])

                    if settings_values:
                        st.caption(
                            f'Besetzung: Früh {settings_values.get("min_f", "–")} · '
                            f'Spät {settings_values.get("min_s", "–")} · '
                            f'Nacht {settings_values.get("min_n", "–")}'
                        )
                    use_col, delete_col = st.columns(2)
                    with use_col:
                        if st.button("Einstellungen übernehmen", key=f'use_template_{template["id"]}', width="stretch"):
                            apply_template_settings(settings_values)
                            st.success("Die Einstellungen wurden übernommen.")
                            st.rerun()
                    with delete_col:
                        if st.button("Vorlage löschen", key=f'delete_template_{template["id"]}', width="stretch"):
                            delete_vorlage(int(template["id"]))
                            st.rerun()

                    if plan_values:
                        st.markdown("**Monatsplan in einen neuen Monat kopieren**")
                        c1, c2 = st.columns(2)
                        today = date.today()
                        with c1:
                            copy_month = st.selectbox(
                                "Zielmonat",
                                list(MONATSNAMEN),
                                index=today.month - 1,
                                format_func=lambda value: MONATSNAMEN[value],
                                key=f'copy_template_month_{template["id"]}',
                            )
                        with c2:
                            copy_year = st.number_input(
                                "Zieljahr",
                                min_value=2020,
                                max_value=2100,
                                value=today.year,
                                step=1,
                                key=f'copy_template_year_{template["id"]}',
                            )
                        if st.button("Monatsplan kopieren", key=f'copy_template_plan_{template["id"]}', type="primary", width="stretch"):
                            st.session_state.dienstplan_vorlage = copy_plan_to_month(
                                plan_values,
                                int(copy_year),
                                int(copy_month),
                                employees,
                                vorgaben,
                            )
                            st.session_state.aktiver_dienstplan_id = None
                            st.session_state.dienstplan_name = f"Dienstplan {MONATSNAMEN[int(copy_month)]} {int(copy_year)}"
                            st.session_state.dienstplan_editor_version += 1
                            st.success("Der Monatsplan wurde kopiert und geöffnet.")
                            st.rerun()


elif page == "Dienstplan":
    page_header(
        "Dienstplan erstellen",
        "Wähle einen Monat, übernimm Abwesenheiten und erstelle den Plan automatisch oder bearbeite ihn selbst.",
    )

    if not employees:
        st.warning("Bitte lege zuerst Mitarbeitende an.")
    else:
        new_tab, copy_tab, load_tab = st.tabs(["✨ Neuer Monat", "📋 Vorherigen Plan kopieren", "📂 Gespeicherten Plan laden"])

        today = date.today()
        with new_tab:
            c1, c2 = st.columns(2)
            with c1:
                selected_month = st.selectbox(
                    "Monat",
                    list(MONATSNAMEN),
                    index=today.month - 1,
                    format_func=lambda value: MONATSNAMEN[value],
                    key="new_plan_month",
                )
            with c2:
                selected_year = st.number_input(
                    "Jahr",
                    min_value=2020,
                    max_value=2100,
                    value=today.year,
                    step=1,
                    key="new_plan_year",
                )
            open_column, auto_column = st.columns(2)
            with open_column:
                if st.button("Leeren Monatsplan öffnen", width="stretch"):
                    st.session_state.dienstplan_vorlage = apply_vorgaben(
                        int(selected_year), int(selected_month), employees, vorgaben
                    )
                    st.session_state.aktiver_dienstplan_id = None
                    st.session_state.dienstplan_name = f"Dienstplan {MONATSNAMEN[int(selected_month)]} {int(selected_year)}"
                    st.session_state.auto_plan_result = None
                    st.session_state.dienstplan_editor_version += 1
                    st.rerun()
            with auto_column:
                if st.button("Automatisch erstellen", type="primary", width="stretch"):
                    automatic_plan, automatic_result = create_automatic_month(
                        int(selected_year),
                        int(selected_month),
                        employees,
                        settings,
                        vorgaben,
                    )
                    st.session_state.dienstplan_vorlage = automatic_plan
                    st.session_state.aktiver_dienstplan_id = None
                    st.session_state.dienstplan_name = f"Dienstplan {MONATSNAMEN[int(selected_month)]} {int(selected_year)}"
                    st.session_state.auto_plan_result = automatic_result
                    st.session_state.dienstplan_editor_version += 1
                    st.rerun()

        with copy_tab:
            if not saved_plans:
                st.info("Noch kein gespeicherter Dienstplan vorhanden.")
            else:
                saved_options = {
                    f'{plan["name"]} · {german_date(plan["startdatum"])}': int(plan["id"])
                    for plan in saved_plans
                }
                source_label = st.selectbox("Welcher Plan soll kopiert werden?", list(saved_options), key="copy_source_plan")
                c1, c2 = st.columns(2)
                with c1:
                    copy_month = st.selectbox(
                        "Zielmonat",
                        list(MONATSNAMEN),
                        index=today.month - 1,
                        format_func=lambda value: MONATSNAMEN[value],
                        key="copy_plan_month",
                    )
                with c2:
                    copy_year = st.number_input(
                        "Zieljahr",
                        min_value=2020,
                        max_value=2100,
                        value=today.year,
                        step=1,
                        key="copy_plan_year",
                    )
                st.caption("Die Dienste werden nach Kalendertag übernommen. Neue Abwesenheiten haben immer Vorrang.")
                if st.button("Plan kopieren und öffnen", type="primary", width="stretch"):
                    source = load_dienstplan(saved_options[source_label])
                    if source:
                        st.session_state.dienstplan_vorlage = copy_plan_to_month(
                            source["plan"], int(copy_year), int(copy_month), employees, vorgaben
                        )
                        st.session_state.aktiver_dienstplan_id = None
                        st.session_state.dienstplan_name = f"Dienstplan {MONATSNAMEN[int(copy_month)]} {int(copy_year)}"
                        st.session_state.auto_plan_result = None
                        st.session_state.dienstplan_editor_version += 1
                        st.rerun()

        with load_tab:
            if not saved_plans:
                st.info("Noch kein Plan gespeichert.")
            else:
                load_options = {
                    f'{plan["name"]} · {german_date(plan["startdatum"])}': int(plan["id"])
                    for plan in saved_plans
                }
                selected_saved = st.selectbox("Gespeicherten Plan auswählen", list(load_options), key="load_saved_plan")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("Plan öffnen", type="primary", width="stretch"):
                        loaded = load_dienstplan(load_options[selected_saved])
                        if loaded:
                            st.session_state.dienstplan_vorlage = normalise_plan(loaded["plan"])
                            st.session_state.aktiver_dienstplan_id = loaded["id"]
                            st.session_state.dienstplan_name = loaded["name"]
                            st.session_state.auto_plan_result = None
                            st.session_state.dienstplan_editor_version += 1
                            st.rerun()
                with c2:
                    if st.button("Gespeicherten Plan löschen", width="stretch"):
                        delete_dienstplan(load_options[selected_saved])
                        st.rerun()

        if not st.session_state.dienstplan_vorlage:
            st.info("Öffne oben einen neuen Monat oder lade einen gespeicherten Plan.")
        else:
            st.divider()
            current_plan = st.session_state.dienstplan_vorlage
            year, month = month_from_plan(current_plan)
            current_errors, current_warnings = validate_plan(current_plan, employees, settings)
            current_stats = calculate_hours(current_plan, employees, settings)
            team_target = float(current_stats["Soll"].sum()) if not current_stats.empty else 0.0
            team_actual = float(current_stats["Ist"].sum()) if not current_stats.empty else 0.0

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                kpi_card("Geöffneter Monat", f"{MONATSNAMEN[month]} {year}", st.session_state.dienstplan_name)
            with c2:
                kpi_card("Geplante Stunden", f"{team_actual:.1f}", f"von {team_target:.1f} Sollstunden")
            with c3:
                kpi_card("Fehler", str(len(current_errors)), "müssen vor Freigabe geprüft werden")
            with c4:
                kpi_card("Hinweise", str(len(current_warnings)), "zum Beispiel Unterbesetzung")

            show_plan_legend()

            auto_result = st.session_state.get("auto_plan_result")
            if auto_result:
                shortages = auto_result.get("shortages", [])
                st.success(
                    f'Automatisch eingetragen: {int(auto_result.get("assigned", 0))} Dienste. '
                    "Gespeicherte Abwesenheiten wurden nicht überschrieben."
                )
                if shortages:
                    with st.expander(f"Offene Besetzungen: {len(shortages)}"):
                        for message in shortages[:50]:
                            st.warning(message)

            edit_plan_tab, readable_tab, check_tab, save_tab = st.tabs(
                ["✏️ Plan bearbeiten", "👀 Lesbare Ansicht", "✅ Prüfung und Stunden", "💾 Speichern und Export"]
            )

            matrix = plan_to_matrix(current_plan, employees, settings)
            day_columns = [column for column in matrix.columns if re.match(r"^\d{2} ", str(column))]
            column_config = {
                "Mitarbeitende": st.column_config.TextColumn("Mitarbeitende", disabled=True, width="medium"),
                "Soll": st.column_config.NumberColumn("Soll", disabled=True, format="%.1f", width="small"),
                "Ist": st.column_config.NumberColumn("Ist", disabled=True, format="%.1f", width="small"),
                "Rest": st.column_config.NumberColumn(
                    "Rest",
                    disabled=True,
                    format="%.1f",
                    width="small",
                    help="Noch verfügbare Stunden bis zum Monatssoll.",
                ),
            }
            for column in day_columns:
                column_config[column] = st.column_config.SelectboxColumn(
                    column,
                    options=DIENST_CODES,
                    required=False,
                    width="small",
                    help="Dienst auswählen oder Feld leer lassen.",
                )

            with edit_plan_tab:
                st.caption("Klicke in ein Tagesfeld und wähle den Dienst. Danach unten auf „Änderungen übernehmen“ klicken.")
                edited_matrix = st.data_editor(
                    matrix,
                    hide_index=True,
                    width="stretch",
                    num_rows="fixed",
                    disabled=["Mitarbeitende", "Soll", "Ist", "Rest"],
                    column_config=column_config,
                    key=f"month_editor_{st.session_state.dienstplan_editor_version}",
                    height=min(760, 105 + len(employees) * 38),
                )
                preview_plan = matrix_to_plan(edited_matrix, current_plan)
                preview_stats = calculate_hours(preview_plan, employees, settings)
                over_target = [
                    row for _, row in preview_stats.iterrows()
                    if float(row["Ist"]) > float(row["Soll"]) + 0.05
                ]
                if st.button("Änderungen übernehmen", type="primary", width="stretch"):
                    if over_target:
                        names = ", ".join(str(row["Mitarbeitende"]) for row in over_target)
                        st.error(f"Noch nicht übernommen: {names} wäre über den Sollstunden.")
                    else:
                        st.session_state.dienstplan_vorlage = preview_plan
                        st.session_state.dienstplan_editor_version += 1
                        st.rerun()

            # data_editor exists even when another tab is selected, therefore preview_plan is available.
            if "preview_plan" not in locals():
                preview_plan = current_plan
                preview_stats = calculate_hours(preview_plan, employees, settings)
                over_target = []

            with readable_tab:
                st.caption("Diese Ansicht ist nur zum Lesen. Die Farben erleichtern die Kontrolle der Dienste.")
                st.dataframe(
                    readable_plan_matrix(plan_to_matrix(preview_plan, employees, settings)),
                    hide_index=True,
                    width="stretch",
                    height=min(760, 105 + len(employees) * 38),
                )

            errors, warnings = validate_plan(preview_plan, employees, settings)
            with check_tab:
                stats_display = preview_stats.rename(
                    columns={
                        "Soll": "Sollstunden",
                        "Ist": "Geplant",
                        "+/-": "Abweichung",
                        "F": "Früh",
                        "S": "Spät",
                        "N": "Nacht",
                        "U": "Urlaub",
                        "W": "Wunschfrei",
                        "FB": "Fortbildung",
                        "K": "Krank",
                    }
                )
                st.dataframe(stats_display, hide_index=True, width="stretch")

                if errors:
                    with st.expander(f"Fehler: {len(errors)}", expanded=True):
                        for message in errors:
                            st.error(message)
                else:
                    st.success("Keine Fehler gefunden.")

                if warnings:
                    with st.expander(f"Hinweise: {len(warnings)}"):
                        for message in warnings[:60]:
                            st.warning(message)
                        if len(warnings) > 60:
                            st.caption(f"Weitere {len(warnings) - 60} Hinweise werden nicht angezeigt.")

            with save_tab:
                plan_name = st.text_input("Name des Dienstplans", value=st.session_state.dienstplan_name)
                if over_target:
                    st.error("Mindestens eine Person liegt über den Sollstunden. Entferne zuerst einen Dienst.")
                if st.button("Dienstplan speichern", type="primary", width="stretch", disabled=bool(over_target)):
                    plan_id = save_dienstplan(
                        name=plan_name.strip() or "Dienstplan",
                        startdatum=preview_plan[0]["Datum"],
                        zeitraum=len(preview_plan),
                        plan=preview_plan,
                        plan_id=st.session_state.aktiver_dienstplan_id,
                    )
                    st.session_state.aktiver_dienstplan_id = plan_id
                    st.session_state.dienstplan_name = plan_name.strip() or "Dienstplan"
                    st.session_state.dienstplan_vorlage = preview_plan
                    st.success("Der Dienstplan wurde gespeichert.")

                export_name = safe_filename(plan_name)
                export_1, export_2 = st.columns(2)
                with export_1:
                    st.download_button(
                        "Als Excel herunterladen",
                        data=create_excel(preview_plan, plan_name, employees, settings),
                        file_name=f"{export_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                        disabled=bool(over_target),
                    )
                with export_2:
                    st.download_button(
                        "Als PDF herunterladen",
                        data=create_pdf(preview_plan, plan_name, employees, settings),
                        file_name=f"{export_name}.pdf",
                        mime="application/pdf",
                        width="stretch",
                        disabled=bool(over_target),
                    )


elif page == "Auswertung":
    page_header(
        "Auswertung",
        "Vergleiche Soll- und Iststunden, Schichtverteilung und Planungsfehler für den geöffneten Dienstplan.",
    )

    if not st.session_state.dienstplan_vorlage:
        st.info("Öffne zuerst einen Dienstplan auf der Seite „Dienstplan“.")
    elif not employees:
        st.info("Es sind keine Mitarbeitenden vorhanden.")
    else:
        plan = st.session_state.dienstplan_vorlage
        stats = calculate_hours(plan, employees, settings)
        errors, warnings = validate_plan(plan, employees, settings)
        target_total = float(stats["Soll"].sum()) if not stats.empty else 0.0
        actual_total = float(stats["Ist"].sum()) if not stats.empty else 0.0
        remaining_total = target_total - actual_total

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            kpi_card("Sollstunden", f"{target_total:.1f}", "gesamtes Team")
        with c2:
            kpi_card("Geplante Stunden", f"{actual_total:.1f}", "inklusive bezahlter Abwesenheit")
        with c3:
            kpi_card("Noch offen", f"{remaining_total:.1f}", "Stunden bis zum Teamsoll")
        with c4:
            kpi_card("Planprüfung", f"{len(errors)} / {len(warnings)}", "Fehler / Hinweise")

        overview_tab, chart_tab, quality_tab = st.tabs(["Stundenübersicht", "Schichtverteilung", "Planqualität"])
        with overview_tab:
            display = stats.rename(
                columns={
                    "Soll": "Sollstunden",
                    "Ist": "Geplant",
                    "+/-": "Abweichung",
                    "F": "Früh",
                    "S": "Spät",
                    "N": "Nacht",
                    "U": "Urlaub",
                    "W": "Wunschfrei",
                    "FB": "Fortbildung",
                    "K": "Krank",
                }
            )
            st.dataframe(display, hide_index=True, width="stretch")

        with chart_tab:
            chart_data = stats.set_index("Mitarbeitende")[["F", "S", "N"]].rename(
                columns={"F": "Frühdienst", "S": "Spätdienst", "N": "Nachtdienst"}
            )
            st.bar_chart(chart_data)
            st.caption("Die Grafik zeigt die Anzahl der geplanten Früh-, Spät- und Nachtdienste je Person.")

        with quality_tab:
            if errors:
                section_title(f"Fehler ({len(errors)})")
                for message in errors:
                    st.error(message)
            else:
                st.success("Keine Fehler gefunden.")
            if warnings:
                section_title(f"Hinweise ({len(warnings)})")
                for message in warnings[:60]:
                    st.warning(message)


elif page == "Planungsregeln":
    page_header(
        "Planungsregeln",
        "Bestimme selbst, welche Regeln die automatische Planung einhalten soll. Jede Änderung gilt nur für dein Konto.",
    )

    custom_rules = eigene_planungsregeln(settings)
    active_basic_count = sum(
        1
        for key in [
            "regel_ruhezeit_aktiv",
            "regel_max_arbeitstage_aktiv",
            "regel_max_nachtdienste_aktiv",
            "regel_max_wochenenden_aktiv",
            "keine_folgewochenenden",
            "regel_sollstunden_aktiv",
            "regel_fachkraft_aktiv",
            "regel_nacht_zuerst_aktiv",
            "regel_nachtblock_aktiv",
            "regel_wochenende_zuerst_aktiv",
            "regel_wochenende_gleich_aktiv",
        ]
        if bool(settings.get(key, True))
    )
    active_custom_count = sum(1 for rule in custom_rules if bool(rule.get("aktiv", True)))

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Aktive Grundregeln", str(active_basic_count), "für Schutz und Planungsreihenfolge")
    with c2:
        kpi_card("Eigene Regeln", str(len(custom_rules)), f"davon {active_custom_count} aktiv")
    with c3:
        kpi_card("Speicherung", "Automatisch", "persönlich im Benutzerkonto")

    basic_tab, custom_tab, overview_tab = st.tabs(
        ["Grundregeln", "Eigene Wenn–Dann-Regeln", "Regelübersicht"]
    )

    with basic_tab:
        section_title(
            "Schutz- und Belastungsregeln",
            "Ein Häkchen aktiviert die Regel. Der Wert daneben bestimmt die Grenze.",
        )
        with st.form("basic_planning_rules_form"):
            left, right = st.columns(2, gap="large")
            with left:
                rest_active = st.checkbox(
                    "Mindestruhezeit einhalten",
                    value=regel_aktiv(settings, "regel_ruhezeit_aktiv"),
                )
                rest_hours = st.number_input(
                    "Mindestruhezeit in Stunden",
                    min_value=0.0,
                    max_value=24.0,
                    value=float(settings.get("ruhezeit", 11.0)),
                    step=0.5,
                )
                max_work_active = st.checkbox(
                    "Arbeitstage hintereinander begrenzen",
                    value=regel_aktiv(settings, "regel_max_arbeitstage_aktiv"),
                )
                max_work_days = st.number_input(
                    "Maximale Arbeitstage hintereinander",
                    min_value=1,
                    max_value=14,
                    value=int(settings.get("max_arbeitstage", 6)),
                    step=1,
                )
                target_active = st.checkbox(
                    "Persönliche Sollstunden nicht überschreiten",
                    value=regel_aktiv(settings, "regel_sollstunden_aktiv"),
                )
            with right:
                max_nights_active = st.checkbox(
                    "Nachtdienste hintereinander begrenzen",
                    value=regel_aktiv(settings, "regel_max_nachtdienste_aktiv"),
                )
                max_night_days = st.number_input(
                    "Maximale Nachtdienste hintereinander",
                    min_value=1,
                    max_value=10,
                    value=int(settings.get("max_nachtdienste", 4)),
                    step=1,
                )
                max_weekends_active = st.checkbox(
                    "Arbeitswochenenden pro Monat begrenzen",
                    value=regel_aktiv(settings, "regel_max_wochenenden_aktiv"),
                )
                max_weekends = st.number_input(
                    "Maximale Arbeitswochenenden pro Monat",
                    min_value=0,
                    max_value=5,
                    value=int(settings.get("max_wochenenden", 2)),
                    step=1,
                )
                no_consecutive_weekends = st.checkbox(
                    "Keine direkt aufeinanderfolgenden Arbeitswochenenden",
                    value=bool(settings.get("keine_folgewochenenden", True)),
                )

            st.divider()
            section_title(
                "Qualität und Reihenfolge",
                "Diese Regeln steuern Besetzung und Ablauf der automatischen Planung.",
            )
            q1, q2 = st.columns(2, gap="large")
            with q1:
                fachkraft_active = st.checkbox(
                    "Mindestzahl an Fachkräften einhalten",
                    value=regel_aktiv(settings, "regel_fachkraft_aktiv"),
                )
                night_first_active = st.checkbox(
                    "Zuerst alle Nachtdienste planen",
                    value=regel_aktiv(settings, "regel_nacht_zuerst_aktiv"),
                )
                night_block_active = st.checkbox(
                    "Begonnene Nachtblöcke möglichst fortsetzen",
                    value=regel_aktiv(settings, "regel_nachtblock_aktiv"),
                )
            with q2:
                weekend_first_active = st.checkbox(
                    "Wochenenden vor den Werktagen planen",
                    value=regel_aktiv(settings, "regel_wochenende_zuerst_aktiv"),
                )
                same_weekend_active = st.checkbox(
                    "Samstag und Sonntag möglichst gleicher Dienst",
                    value=regel_aktiv(settings, "regel_wochenende_gleich_aktiv"),
                )

            save_basic_rules = st.form_submit_button(
                "Grundregeln speichern", type="primary", width="stretch"
            )

        if save_basic_rules:
            new_settings = dict(settings)
            new_settings.update(
                {
                    "regel_ruhezeit_aktiv": bool(rest_active),
                    "ruhezeit": float(rest_hours),
                    "regel_max_arbeitstage_aktiv": bool(max_work_active),
                    "max_arbeitstage": int(max_work_days),
                    "regel_sollstunden_aktiv": bool(target_active),
                    "regel_max_nachtdienste_aktiv": bool(max_nights_active),
                    "max_nachtdienste": int(max_night_days),
                    "regel_max_wochenenden_aktiv": bool(max_weekends_active),
                    "max_wochenenden": int(max_weekends),
                    "keine_folgewochenenden": bool(no_consecutive_weekends),
                    "regel_fachkraft_aktiv": bool(fachkraft_active),
                    "regel_nacht_zuerst_aktiv": bool(night_first_active),
                    "regel_nachtblock_aktiv": bool(night_block_active),
                    "regel_wochenende_zuerst_aktiv": bool(weekend_first_active),
                    "regel_wochenende_gleich_aktiv": bool(same_weekend_active),
                }
            )
            save_einstellungen(new_settings)
            st.session_state.settings = new_settings
            st.success("Die Grundregeln wurden gespeichert.")
            st.rerun()

        st.info(
            "Die Mindestbesetzung pro Schicht wird weiterhin unter „Einstellungen → Mindestbesetzung“ festgelegt. "
            "Tarifvertrag, Betriebsvereinbarung und interne Vorgaben müssen zusätzlich geprüft werden."
        )

    with custom_tab:
        section_title(
            "Eigene Regel erstellen oder bearbeiten",
            "Wähle eine Regelart aus und fülle die angezeigten Auswahlfelder aus. Programmieren ist nicht nötig.",
        )
        if not employees:
            st.warning("Lege zuerst mindestens eine mitarbeitende Person an.")
        else:
            rule_type_labels = {
                "nicht_gemeinsam": "Zwei Personen nicht gemeinsam einplanen",
                "frei_nach_dienst": "Freie Tage nach einer Dienstart",
                "max_dienstfolge": "Dienstart hintereinander begrenzen",
            }
            edit_options = {"Neue Regel erstellen": "NEW"}
            for rule in custom_rules:
                edit_options[f'Bearbeiten: {rule.get("name", "Eigene Regel")}'] = str(rule.get("id"))

            selected_label = st.selectbox(
                "Neue oder vorhandene Regel",
                list(edit_options),
                key="custom_rule_editor_selection",
            )
            selected_rule_id = edit_options[selected_label]
            selected_rule = next(
                (rule for rule in custom_rules if str(rule.get("id")) == selected_rule_id),
                None,
            )
            selected_type = str((selected_rule or {}).get("typ", "nicht_gemeinsam"))
            type_names = list(rule_type_labels)
            rule_type = st.selectbox(
                "Regelart",
                type_names,
                index=type_names.index(selected_type) if selected_type in type_names else 0,
                format_func=lambda value: rule_type_labels[value],
                key=f"custom_rule_type_{selected_rule_id}",
            )

            employee_ids = [int(person["id"]) for person in employees]
            employee_by_id_ui = {int(person["id"]): person for person in employees}
            shift_labels = {"F": "Frühdienst", "S": "Spätdienst", "N": "Nachtdienst"}
            editor_token = f"{selected_rule_id}_{rule_type}"

            with st.form(f"custom_rule_form_{editor_token}"):
                rule_name = st.text_input(
                    "Name der Regel",
                    value=str((selected_rule or {}).get("name", "")),
                    placeholder="Zum Beispiel: Ilona und Marion nicht zusammen",
                    max_chars=100,
                )
                rule_enabled = st.checkbox(
                    "Regel sofort aktivieren",
                    value=bool((selected_rule or {}).get("aktiv", True)),
                )

                if rule_type == "nicht_gemeinsam":
                    first_default = regel_zahl(selected_rule or {}, "mitarbeitende_a_id", employee_ids[0])
                    second_default = regel_zahl(
                        selected_rule or {},
                        "mitarbeitende_b_id",
                        employee_ids[1] if len(employee_ids) > 1 else employee_ids[0],
                    )
                    first_id = st.selectbox(
                        "Erste Person",
                        employee_ids,
                        index=employee_ids.index(first_default) if first_default in employee_ids else 0,
                        format_func=lambda value: employee_by_id_ui[value]["name"],
                    )
                    second_id = st.selectbox(
                        "Zweite Person",
                        employee_ids,
                        index=employee_ids.index(second_default) if second_default in employee_ids else 0,
                        format_func=lambda value: employee_by_id_ui[value]["name"],
                    )
                    together_shift_options = ["ALL", "F", "S", "N"]
                    together_shift_labels = {"ALL": "Alle Dienstarten", **shift_labels}
                    old_shift = str((selected_rule or {}).get("dienst", "ALL"))
                    selected_shift = st.selectbox(
                        "Gilt für",
                        together_shift_options,
                        index=together_shift_options.index(old_shift) if old_shift in together_shift_options else 0,
                        format_func=lambda value: together_shift_labels[value],
                    )

                elif rule_type == "frei_nach_dienst":
                    scope_ids = [0] + employee_ids
                    scope_default = regel_zahl(selected_rule or {}, "mitarbeitende_id", 0)
                    selected_employee_id = st.selectbox(
                        "Gilt für",
                        scope_ids,
                        index=scope_ids.index(scope_default) if scope_default in scope_ids else 0,
                        format_func=lambda value: "Alle Mitarbeitenden" if value == 0 else employee_by_id_ui[value]["name"],
                    )
                    old_shift = str((selected_rule or {}).get("dienst", "N"))
                    selected_shift = st.selectbox(
                        "Nach welcher Dienstart?",
                        list(shift_labels),
                        index=list(shift_labels).index(old_shift) if old_shift in shift_labels else 2,
                        format_func=lambda value: shift_labels[value],
                    )
                    free_days_value = st.number_input(
                        "Mindestens freie Kalendertage danach",
                        min_value=1,
                        max_value=7,
                        value=max(1, regel_zahl(selected_rule or {}, "freie_tage", 2)),
                        step=1,
                    )

                else:
                    scope_ids = [0] + employee_ids
                    scope_default = regel_zahl(selected_rule or {}, "mitarbeitende_id", 0)
                    selected_employee_id = st.selectbox(
                        "Gilt für",
                        scope_ids,
                        index=scope_ids.index(scope_default) if scope_default in scope_ids else 0,
                        format_func=lambda value: "Alle Mitarbeitenden" if value == 0 else employee_by_id_ui[value]["name"],
                    )
                    old_shift = str((selected_rule or {}).get("dienst", "N"))
                    selected_shift = st.selectbox(
                        "Welche Dienstart?",
                        list(shift_labels),
                        index=list(shift_labels).index(old_shift) if old_shift in shift_labels else 2,
                        format_func=lambda value: shift_labels[value],
                    )
                    maximum_value = st.number_input(
                        "Höchstens hintereinander",
                        min_value=1,
                        max_value=10,
                        value=max(1, regel_zahl(selected_rule or {}, "maximum", 3)),
                        step=1,
                    )

                save_custom_rule = st.form_submit_button(
                    "Regel speichern", type="primary", width="stretch",
                    disabled=rule_type == "nicht_gemeinsam" and len(employee_ids) < 2,
                )

            if save_custom_rule:
                validation_message = ""
                if not rule_name.strip():
                    validation_message = "Bitte gib der Regel einen Namen."
                elif rule_type == "nicht_gemeinsam" and first_id == second_id:
                    validation_message = "Bitte wähle zwei unterschiedliche Personen."

                if validation_message:
                    st.warning(validation_message)
                else:
                    saved_rule = {
                        "id": str((selected_rule or {}).get("id") or uuid4()),
                        "name": rule_name.strip(),
                        "typ": rule_type,
                        "aktiv": bool(rule_enabled),
                    }
                    if rule_type == "nicht_gemeinsam":
                        saved_rule.update(
                            {
                                "mitarbeitende_a_id": int(first_id),
                                "mitarbeitende_b_id": int(second_id),
                                "dienst": selected_shift,
                            }
                        )
                    elif rule_type == "frei_nach_dienst":
                        saved_rule.update(
                            {
                                "mitarbeitende_id": int(selected_employee_id),
                                "dienst": selected_shift,
                                "freie_tage": int(free_days_value),
                            }
                        )
                    else:
                        saved_rule.update(
                            {
                                "mitarbeitende_id": int(selected_employee_id),
                                "dienst": selected_shift,
                                "maximum": int(maximum_value),
                            }
                        )

                    updated_rules = [
                        rule
                        for rule in custom_rules
                        if str(rule.get("id")) != str(saved_rule["id"])
                    ]
                    updated_rules.append(saved_rule)
                    new_settings = {**settings, "eigene_planungsregeln": updated_rules}
                    save_einstellungen(new_settings)
                    st.session_state.settings = new_settings
                    st.success("Die eigene Regel wurde gespeichert.")
                    st.rerun()

            st.divider()
            section_title(
                "Gespeicherte eigene Regeln",
                "Deaktivierte Regeln bleiben gespeichert, werden aber nicht angewendet.",
            )
            if not custom_rules:
                st.info("Noch keine eigene Regel gespeichert.")
            else:
                for rule in custom_rules:
                    rule_id = str(rule.get("id"))
                    enabled = bool(rule.get("aktiv", True))
                    status_icon = "🟢" if enabled else "⚪"
                    with st.expander(f'{status_icon} {rule.get("name", "Eigene Regel")}'):
                        st.write(regel_beschreibung(rule, employees))
                        st.caption("Aktiv" if enabled else "Inaktiv")
                        action_col, delete_col = st.columns(2)
                        with action_col:
                            action_label = "Deaktivieren" if enabled else "Aktivieren"
                            if st.button(action_label, key=f"toggle_rule_{rule_id}", width="stretch"):
                                updated_rules = []
                                for item in custom_rules:
                                    updated = dict(item)
                                    if str(updated.get("id")) == rule_id:
                                        updated["aktiv"] = not enabled
                                    updated_rules.append(updated)
                                new_settings = {**settings, "eigene_planungsregeln": updated_rules}
                                save_einstellungen(new_settings)
                                st.session_state.settings = new_settings
                                st.rerun()
                        with delete_col:
                            if st.button("Löschen", key=f"delete_rule_{rule_id}", width="stretch"):
                                updated_rules = [
                                    item for item in custom_rules
                                    if str(item.get("id")) != rule_id
                                ]
                                new_settings = {**settings, "eigene_planungsregeln": updated_rules}
                                save_einstellungen(new_settings)
                                st.session_state.settings = new_settings
                                st.rerun()

    with overview_tab:
        section_title(
            "Alle Regeln auf einen Blick",
            "Nur aktive Regeln werden bei der automatischen Planung und Planprüfung angewendet.",
        )
        basic_rows = [
            ("Mindestruhezeit", "regel_ruhezeit_aktiv", f'{float(settings.get("ruhezeit", 11.0)):g} Stunden'),
            ("Maximale Arbeitstage", "regel_max_arbeitstage_aktiv", str(int(settings.get("max_arbeitstage", 6)))),
            ("Maximale Nachtdienste", "regel_max_nachtdienste_aktiv", str(int(settings.get("max_nachtdienste", 4)))),
            ("Maximale Arbeitswochenenden", "regel_max_wochenenden_aktiv", str(int(settings.get("max_wochenenden", 2)))),
            ("Keine Folgewochenenden", "keine_folgewochenenden", "–"),
            ("Sollstunden nicht überschreiten", "regel_sollstunden_aktiv", "–"),
            ("Fachkraftbesetzung", "regel_fachkraft_aktiv", "–"),
            ("Nachtdienste zuerst", "regel_nacht_zuerst_aktiv", "–"),
            ("Nachtblöcke fortsetzen", "regel_nachtblock_aktiv", "–"),
            ("Wochenenden zuerst", "regel_wochenende_zuerst_aktiv", "–"),
            ("Gleicher Wochenenddienst", "regel_wochenende_gleich_aktiv", "–"),
        ]
        overview_rows = [
            {
                "Regel": name,
                "Art": "Grundregel",
                "Status": "Aktiv" if bool(settings.get(key, True)) else "Inaktiv",
                "Wert": value,
            }
            for name, key, value in basic_rows
        ]
        overview_rows.extend(
            {
                "Regel": str(rule.get("name", "Eigene Regel")),
                "Art": "Eigene Regel",
                "Status": "Aktiv" if bool(rule.get("aktiv", True)) else "Inaktiv",
                "Wert": regel_beschreibung(rule, employees),
            }
            for rule in custom_rules
        )
        st.dataframe(
            pd.DataFrame(overview_rows),
            hide_index=True,
            width="stretch",
            column_config={
                "Regel": st.column_config.TextColumn("Regel", width="medium"),
                "Art": st.column_config.TextColumn("Art", width="small"),
                "Status": st.column_config.TextColumn("Status", width="small"),
                "Wert": st.column_config.TextColumn("Wert / Beschreibung", width="large"),
            },
        )


elif page == "Feedback & Hilfe":
    page_header(
        "Kontakt",
        "Hast du eine Frage, eine Idee oder ein Problem? Schreib uns einfach eine Nachricht.",
    )

    try:
        get_feedback(limit=1)
        feedback_available = True
    except Exception:
        feedback_available = False

    if not feedback_available:
        st.warning(
            "Die Feedback-Speicherung ist noch nicht aktiviert. Führe zuerst einmal "
            "die Datei „feedback_setup.sql“ im Supabase SQL Editor aus."
        )

    _, contact_column, _ = st.columns([1, 1.6, 1])
    with contact_column:
        with st.container(border=True):
            section_title(
                "Nachricht senden",
                "Wir lesen jede Nachricht und nutzen sie, um NursePlan Pro zu verbessern.",
            )
            with st.form("feedback_form", clear_on_submit=True):
                feedback_subject = st.text_input(
                    "Betreff (optional)",
                    placeholder="Worum geht es?",
                    max_chars=120,
                )
                feedback_message = st.text_area(
                    "Deine Nachricht",
                    placeholder="Schreib uns kurz, was du mitteilen möchtest …",
                    height=180,
                    max_chars=3000,
                )
                st.caption(
                    "Bitte keine Patientendaten, Passwörter oder Zugangsschlüssel eintragen."
                )
                submit_feedback = st.form_submit_button(
                    "Nachricht senden",
                    type="primary",
                    width="stretch",
                    disabled=not feedback_available,
                )

            if submit_feedback:
                if len(feedback_message.strip()) < 10:
                    st.warning("Bitte schreibe eine kurze Nachricht mit mindestens 10 Zeichen.")
                else:
                    try:
                        add_feedback(
                            kategorie="Frage",
                            bewertung=5,
                            betreff=feedback_subject.strip() or "Allgemeine Rückmeldung",
                            nachricht=feedback_message,
                            kontakt_erlaubt=True,
                            kontakt_email=current_user_email(),
                        )
                        st.success("Danke! Deine Nachricht wurde gesendet.")
                    except Exception:
                        st.error(
                            "Die Nachricht konnte gerade nicht gesendet werden. "
                            "Bitte versuche es später noch einmal."
                        )

        st.caption(
            "🔒 Deine Nachricht ist privat. Wenn eine Antwort nötig ist, nutzen wir die "
            "E-Mail-Adresse deines Kontos."
        )


elif page == "Einstellungen":
    page_header(
        "Einstellungen",
        "Lege Arbeitszeiten, Mindestbesetzung und Planungsgrenzen fest. Diese Werte gelten für neue automatische Dienstpläne.",
    )

    general_tab, shifts_tab, staffing_tab = st.tabs(
        ["Allgemein", "Arbeitszeiten", "Mindestbesetzung"]
    )

    with general_tab:
        section_title("Monatssoll", "Dieser Wert dient zur Berechnung des Beschäftigungsumfangs.")
        fulltime_target = st.number_input(
            "Vollzeit-Sollstunden im Monat",
            min_value=1.0,
            max_value=300.0,
            value=float(settings["vollzeit_monatssoll"]),
            step=0.5,
        )

    shift_values = {}
    with shifts_tab:
        section_title("Arbeitszeiten", "Beginn, Ende und bezahlte Stunden werden für Ruhezeit und Stundenkonto benötigt.")
        for code, prefix, title in [
            ("F", "f", "Frühdienst"),
            ("S", "s", "Spätdienst"),
            ("N", "n", "Nachtdienst"),
        ]:
            st.markdown(f"### {code} · {title}")
            c1, c2, c3 = st.columns(3)
            with c1:
                start_value = st.time_input(
                    "Beginn",
                    value=parse_clock(settings[f"{prefix}_start"]),
                    key=f"settings_{prefix}_start_new",
                )
            with c2:
                end_value = st.time_input(
                    "Ende",
                    value=parse_clock(settings[f"{prefix}_ende"]),
                    key=f"settings_{prefix}_end_new",
                )
            with c3:
                hours_value = st.number_input(
                    "Bezahlte Stunden",
                    min_value=0.0,
                    max_value=24.0,
                    value=float(settings[f"{prefix}_stunden"]),
                    step=0.1,
                    key=f"settings_{prefix}_hours_new",
                )
            shift_values[prefix] = {
                "start": format_clock(start_value),
                "end": format_clock(end_value),
                "hours": float(hours_value),
            }
            st.divider()

    staffing_values = {}
    with staffing_tab:
        section_title("Mindestbesetzung je Dienst", "Die Automatik versucht diese Besetzung an jedem Kalendertag zu erreichen.")
        for code, prefix, title in [
            ("F", "f", "Frühdienst"),
            ("S", "s", "Spätdienst"),
            ("N", "n", "Nachtdienst"),
        ]:
            c1, c2 = st.columns(2)
            with c1:
                minimum_value = st.number_input(
                    f"{title}: Personen insgesamt",
                    min_value=0,
                    max_value=30,
                    value=int(settings[f"min_{prefix}"]),
                    step=1,
                    key=f"settings_{prefix}_minimum_new",
                )
            with c2:
                fach_value = st.number_input(
                    f"{title}: davon Fachkräfte",
                    min_value=0,
                    max_value=30,
                    value=int(settings.get(f"min_fach_{prefix}", 1)),
                    step=1,
                    key=f"settings_{prefix}_fach_new",
                )
            staffing_values[prefix] = {"minimum": int(minimum_value), "fach": int(fach_value)}

    if st.button("Alle Einstellungen speichern", type="primary", width="stretch"):
        new_settings = dict(settings)
        new_settings["vollzeit_monatssoll"] = float(fulltime_target)
        new_settings["max_plusstunden"] = 0.0
        for prefix, values in shift_values.items():
            new_settings[f"{prefix}_start"] = values["start"]
            new_settings[f"{prefix}_ende"] = values["end"]
            new_settings[f"{prefix}_stunden"] = values["hours"]
        for prefix, values in staffing_values.items():
            new_settings[f"min_{prefix}"] = values["minimum"]
            new_settings[f"min_fach_{prefix}"] = values["fach"]

        save_einstellungen(new_settings)
        st.session_state.settings = new_settings
        st.success("Die Einstellungen wurden gespeichert.")
        st.rerun()
