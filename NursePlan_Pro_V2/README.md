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

## Intelligente Dienstplanung V2

Neben der klassischen Automatik enthält NursePlan Pro eine globale Planung mit
**Google OR-Tools (CP-SAT)**. In der App steht sie über den Button
**„✨ Intelligent planen (V2)“** zur Verfügung.

V2 berücksichtigt gemeinsam:

- Mindestbesetzung und erforderliche Fachkräfte je Schicht
- vollständige Abwesenheitszeiträume inklusive dazwischenliegender Wochenenden
- Sollstunden, Ruhezeit sowie maximale Arbeits- und Nachtdienstfolgen
- zusammenhängende Nachtblöcke und freie Tage nach dem Block
- maximale, aufeinanderfolgende und möglichst einheitlich besetzte Arbeitswochenenden
- ruhige Früh-/Spät-Dienstfolgen und eine proportionale Stundenverteilung
- aktive eigene Wenn-Dann-Regeln

Die Optimierung arbeitet in festen Prioritätsstufen: zuerst Besetzung und
Fachkräfte, danach Nachtblöcke und Wochenenden und zuletzt Sollstunden und
Fairness. Eine wichtigere Stufe wird durch eine spätere Optimierung nicht wieder
verschlechtert. Der Solver verwendet weiterhin bewusst nur einen Worker.

Wenn nicht alle Anforderungen gleichzeitig erfüllbar sind, erstellt V2 den
bestmöglichen Plan und weist offene Besetzungen sichtbar aus. Die klassische
Planung bleibt als Alternative erhalten.

## Wichtiger Beta-Hinweis

Vor einer öffentlichen Nutzung mit echten Beschäftigtendaten müssen Impressum,
Datenschutzerklärung, Auftragsverarbeitung, Löschkonzept und die betrieblichen
Regeln fachlich und rechtlich geprüft werden.
