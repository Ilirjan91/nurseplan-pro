import sqlite3
from datetime import datetime
from pathlib import Path


DATABASE_PATH = Path(__file__).resolve().parent / "nurseplan.db"


def get_connection():
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def create_tables():
    with get_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS mitarbeitende (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                stellenanteil TEXT NOT NULL,
                wochenstunden REAL NOT NULL,
                qualifikation TEXT NOT NULL,
                erfahrung TEXT NOT NULL,
                nachtdienst TEXT NOT NULL
            )
            """
        )

        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS dienstplan (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum_iso TEXT NOT NULL UNIQUE,
                datum TEXT NOT NULL,
                wochentag TEXT NOT NULL,
                fruehdienst TEXT NOT NULL DEFAULT '',
                mitteldienst TEXT NOT NULL DEFAULT '',
                spaetdienst TEXT NOT NULL DEFAULT '',
                nachtdienst TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT ''
            )
            """
        )


def add_mitarbeitende(
    name,
    stellenanteil,
    wochenstunden,
    qualifikation,
    erfahrung,
    nachtdienst,
):
    with get_connection() as connection:
        connection.execute(
            """
            INSERT INTO mitarbeitende (
                name,
                stellenanteil,
                wochenstunden,
                qualifikation,
                erfahrung,
                nachtdienst
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                name,
                stellenanteil,
                wochenstunden,
                qualifikation,
                erfahrung,
                nachtdienst,
            ),
        )


def get_mitarbeitende():
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT
                id,
                name,
                stellenanteil,
                wochenstunden,
                qualifikation,
                erfahrung,
                nachtdienst
            FROM mitarbeitende
            ORDER BY name COLLATE NOCASE
            """
        ).fetchall()

    return [dict(row) for row in rows]


def delete_mitarbeitende(mitarbeitende_id):
    with get_connection() as connection:
        connection.execute(
            "DELETE FROM mitarbeitende WHERE id = ?",
            (mitarbeitende_id,),
        )


def save_dienstplan(plan_rows):
    """
    Speichert den aktuellen Plan dauerhaft.
    Bereits vorhandene Einträge mit demselben Datum werden aktualisiert.
    """
    if not plan_rows:
        return

    with get_connection() as connection:
        for row in plan_rows:
            datum_anzeige = str(row.get("Datum", "")).strip()

            try:
                datum_iso = datetime.strptime(
                    datum_anzeige,
                    "%d.%m.%Y",
                ).date().isoformat()
            except ValueError as exc:
                raise ValueError(
                    f"Ungültiges Datum im Dienstplan: {datum_anzeige}"
                ) from exc

            connection.execute(
                """
                INSERT INTO dienstplan (
                    datum_iso,
                    datum,
                    wochentag,
                    fruehdienst,
                    mitteldienst,
                    spaetdienst,
                    nachtdienst,
                    status
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(datum_iso) DO UPDATE SET
                    datum = excluded.datum,
                    wochentag = excluded.wochentag,
                    fruehdienst = excluded.fruehdienst,
                    mitteldienst = excluded.mitteldienst,
                    spaetdienst = excluded.spaetdienst,
                    nachtdienst = excluded.nachtdienst,
                    status = excluded.status
                """,
                (
                    datum_iso,
                    datum_anzeige,
                    str(row.get("Wochentag", "")),
                    str(row.get("Frühdienst", "")),
                    str(row.get("Mitteldienst", "")),
                    str(row.get("Spätdienst", "")),
                    str(row.get("Nachtdienst", "")),
                    str(row.get("Status", "")),
                ),
            )


def get_dienstplan():
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT
                datum,
                wochentag,
                fruehdienst,
                mitteldienst,
                spaetdienst,
                nachtdienst,
                status
            FROM dienstplan
            ORDER BY datum_iso
            """
        ).fetchall()

    plan = []

    for index, row in enumerate(rows, start=1):
        plan.append(
            {
                "Tag": index,
                "Datum": row["datum"],
                "Wochentag": row["wochentag"],
                "Frühdienst": row["fruehdienst"],
                "Mitteldienst": row["mitteldienst"],
                "Spätdienst": row["spaetdienst"],
                "Nachtdienst": row["nachtdienst"],
                "Status": row["status"],
            }
        )

    return plan
