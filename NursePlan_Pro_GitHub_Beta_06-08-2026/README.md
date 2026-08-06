# NursePlan Pro

NursePlan Pro ist eine Streamlit-Anwendung zur Dienstplanung für Pflegeteams.
Konten, Mitarbeitende, Abwesenheiten, Einstellungen, Regeln, Dienstpläne und
Feedback werden getrennt je Benutzerkonto in Supabase gespeichert.

## Online-Beta bereitstellen

1. Dieses Projekt in ein privates GitHub-Repository hochladen.
2. `.streamlit/secrets.toml` niemals hochladen. Sie ist bereits in `.gitignore` ausgeschlossen.
3. In Streamlit Community Cloud eine neue App aus dem Repository erstellen.
4. Als Startdatei `app.py` auswählen und Python 3.12 verwenden.
5. In den Streamlit-Secrets diese Werte eintragen:

```toml
SUPABASE_URL = "https://DEIN-PROJEKT.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "DEIN-PUBLISHABLE-KEY"
APP_URL = "https://DEINE-APP.streamlit.app"
```

Nur den Publishable Key verwenden. Niemals Secret Key, `service_role`,
Datenbankpasswort oder andere private Zugangsdaten in GitHub speichern.

6. Nach der Bereitstellung die feste HTTPS-App-Adresse in Supabase unter
   **Authentication → URL Configuration** als Site URL und Redirect URL ergänzen.
7. Registrierung, Anmeldung, Passwort-Zurücksetzen, Feedback und Datentrennung
   mit zwei Testkonten prüfen.

## Lokal starten

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Weitere Einrichtungsschritte stehen in `ANLEITUNG.md`.

## Wichtiger Beta-Hinweis

Vor einer öffentlichen Nutzung mit echten Beschäftigtendaten müssen Impressum,
Datenschutzerklärung, Auftragsverarbeitung, Löschkonzept und die betrieblichen
Regeln fachlich und rechtlich geprüft werden.
