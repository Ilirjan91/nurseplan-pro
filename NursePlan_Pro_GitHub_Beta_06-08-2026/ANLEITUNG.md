# NursePlan Pro mit Supabase-Login

Diese Version enthält:

- Registrierung mit Name, E-Mail-Adresse und Passwort
- Anmeldung und Abmeldung
- Option **„Auf diesem Gerät angemeldet bleiben“** mit automatischer Wiederanmeldung nach einem Browser-Refresh
- Einfache Seite **„Kontakt“** für Fragen, Ideen und Probleme
- Supabase statt lokaler SQLite-Datenbank
- getrennte Daten für jedes Konto durch Row Level Security (RLS)
- eine eigene Seite **Planungsregeln** mit aktivierbaren Grundregeln
- einen einfachen Wenn–Dann-Editor für persönliche Regeln
- den verbesserten Dienstplan-Algorithmus: auf Wunsch zuerst Nachtdienste, dann Wochenenden, danach Werktage

## Planungsregeln verwenden

1. Öffne links **Planungsregeln**.
2. Aktiviere unter **Grundregeln** nur die Regeln, die für deinen Bereich gelten.
3. Öffne **Eigene Wenn–Dann-Regeln**, um zum Beispiel zwei Personen nicht gemeinsam einzuplanen oder freie Tage nach einem Nachtdienst festzulegen.
4. Speichere die Regel und prüfe anschließend unter **Regelübersicht** ihren Status.
5. Erstelle den Dienstplan neu oder öffne **Prüfung und Stunden**, damit die aktiven Regeln angewendet werden.

Die Regeln werden im jeweiligen Benutzerkonto gespeichert. Es ist keine zusätzliche SQL-Tabelle notwendig. Gesetzliche, tarifliche und betriebliche Vorgaben müssen trotzdem fachlich geprüft werden.

## Wichtig vor dem Start

Deine bisherige Datei `nurseplan.db` wird nicht automatisch nach Supabase übertragen. Die alte Datei bitte aufbewahren. Eine Datenübertragung kann nach dem erfolgreichen Login als eigener Schritt erfolgen.

## 1. Supabase-Werte lokal eintragen

1. Öffne in Supabase dein Projekt **NursePlan Pro**.
2. Öffne **Project Settings → API Keys**.
3. Kopiere die **Project URL** und den **Publishable Key**. Wenn dein Projekt nur `anon public` anzeigt, darf dieser Wert stattdessen verwendet werden.
4. Öffne im Projektordner den Unterordner `.streamlit`.
5. Kopiere `secrets.example.toml` und nenne die Kopie `secrets.toml`.
6. Trage die beiden Werte ausschließlich in deiner lokalen `secrets.toml` ein.

So sieht die Datei aus:

```toml
SUPABASE_URL = "https://DEIN-PROJEKT.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "DEIN-PUBLISHABLE-KEY"
```

Niemals `service_role`, `sb_secret_...`, das Datenbankpasswort oder andere geheime Schlüssel verwenden oder versenden.

## 2. Pakete installieren

PowerShell im Projektordner öffnen und ausführen:

```powershell
python -m pip install -r requirements.txt
```

Nach einem Update bitte diesen Befehl erneut ausführen. Dadurch wird auch die Komponente installiert, die die geschützte Browser-Sitzung verwaltet.

## 3. App starten

```powershell
streamlit run app.py
```

Danach öffnet sich die App normalerweise unter `http://localhost:8501`.

## 4. Erstes Konto testen

1. In NursePlan Pro auf **Konto erstellen** klicken.
2. Name, E-Mail-Adresse und ein Passwort mit mindestens 8 Zeichen eingeben.
3. Falls Supabase eine Bestätigungs-E-Mail sendet, den Link öffnen.
4. Danach in NursePlan Pro anmelden.
5. Eine Testperson anlegen und prüfen, ob sie gespeichert wird.

## 5. Anmeldung nach einem Refresh testen

1. Beim Anmelden **„Auf diesem Gerät angemeldet bleiben“** aktivieren.
2. Die Seite mit `Strg + F5` aktualisieren.
3. NursePlan Pro muss das Konto automatisch wiederherstellen.
4. Auf **Abmelden** klicken und die Seite erneut aktualisieren.
5. Danach muss wieder die Anmeldeseite erscheinen.

Die gespeicherte Sitzung gilt höchstens 30 Tage. Diese Option nur auf einem persönlichen oder geschützten Gerät verwenden.

## 6. Feedback-Funktion aktivieren

1. Öffne in Supabase den **SQL Editor**.
2. Öffne aus dem NursePlan-Pro-Ordner die Datei `feedback_setup.sql`.
3. Kopiere den gesamten Inhalt in den SQL Editor und führe ihn einmal aus.
4. Öffne danach in NursePlan Pro links **Kontakt**.
5. Sende eine Testnachricht und prüfe sie im Supabase **Table Editor → feedback**.

Die SQL-Datei legt nur die neue Feedback-Tabelle mit sicheren RLS-Regeln an. Vorhandene Tabellen und Daten werden nicht gelöscht. Nutzer dürfen ausschließlich eigene Rückmeldungen senden und sehen. Die Verwaltung aller Rückmeldungen erfolgt im Supabase-Dashboard.

## 7. „Passwort vergessen“ aktivieren

1. Öffne in Supabase **Authentication → URL Configuration**.
2. Trage unter **Redirect URLs** diese Adresse ein:

   `http://localhost:8501/**`

3. Speichere die Einstellung.
4. Öffne NursePlan Pro und klicke bei der Anmeldung auf **Passwort vergessen?**.
5. Fordere den Link an, öffne die E-Mail im selben Browser und lege ein neues Passwort fest.

Lokal verwendet NursePlan Pro automatisch `http://localhost:8501`. Bei der späteren Veröffentlichung muss `APP_URL = "https://DEINE-APP-ADRESSE"` in den Secrets ergänzt und diese HTTPS-Adresse zusätzlich bei den Supabase Redirect URLs eingetragen werden. Das neue Passwort wird direkt vom Browser an Supabase Auth übertragen und nicht in NursePlan Pro gespeichert.

## 8. Datenschutz prüfen

Vor der Veröffentlichung mit zwei verschiedenen Testkonten prüfen:

1. Mit Konto A eine Testperson speichern.
2. Abmelden und mit Konto B anmelden.
3. Konto B darf die Testperson von Konto A nicht sehen.
4. Wieder mit Konto A anmelden und prüfen, ob die Testperson weiterhin vorhanden ist.

Vor einem öffentlichen Start werden außerdem mindestens Impressum, Datenschutzerklärung, Passwort-Zurücksetzen und eine klare Regelung für echte Beschäftigtendaten benötigt.

## 9. Online-Beta mit Streamlit Community Cloud

1. Lade die Dateien aus dem vorbereiteten GitHub-Paket in ein privates Repository hoch.
2. Prüfe vor dem Upload noch einmal, dass `secrets.toml` und `nurseplan.db` nicht enthalten sind.
3. Öffne Streamlit Community Cloud und erstelle eine App aus diesem Repository.
4. Wähle `app.py` als Startdatei und Python 3.12.
5. Trage Project URL, Publishable Key und die neue `APP_URL` ausschließlich in den Streamlit-Secrets ein.
6. Ergänze die feste HTTPS-Adresse anschließend in Supabase unter **Authentication → URL Configuration**.
7. Teste die Online-App zuerst mit zwei Testkonten und ohne echte Beschäftigtendaten.

Die vollständigen Online-Schritte und das benötigte Secrets-Beispiel stehen in `README.md`.
