"""Intelligente Monatsplanung fuer NursePlan Pro mit Google OR-Tools.

Das Modul ist bewusst unabhaengig von Streamlit und Supabase. Dadurch kann der
Solver separat getestet werden und erhaelt nur die bereits geladenen Daten der
aktuellen Benutzersitzung.
"""

from __future__ import annotations

import calendar
import re
from copy import deepcopy
from datetime import date, datetime, timedelta
from typing import Any


SHIFT_CODES = ("F", "S", "N")
CODE_TO_COLUMN = {
    "F": "Frühdienst",
    "S": "Spätdienst",
    "N": "Nachtdienst",
}
ABSENCE_COLUMNS = ("Urlaub", "Wunschfrei", "Fortbildung", "Krank", "Frei")
PAID_ABSENCE_COLUMNS = ("Urlaub", "Fortbildung", "Krank")
FACHKRAFT_QUALIFIKATIONEN = {
    "Pflegefachkraft",
    "Praxisanleitung",
    "Stroke Nurse",
    "Stationsleitung",
}
ALLOWED_SHIFT_CODES = {
    "Alle Dienste": {"F", "S", "N"},
    "Nur Frühdienst": {"F"},
    "Nur Spätdienst": {"S"},
    "Nur Nachtdienst": {"N"},
}
HOUR_SCALE = 100


class OrToolsUnavailableError(RuntimeError):
    """OR-Tools ist in der aktuellen Python-Umgebung nicht installiert."""


def _cp_model_module():
    try:
        from ortools.sat.python import cp_model
    except ImportError as exc:  # pragma: no cover - nur bei kaputter Installation
        raise OrToolsUnavailableError(
            "Google OR-Tools fehlt. Bitte requirements.txt installieren."
        ) from exc
    return cp_model


def _split_names(value: Any) -> list[str]:
    if value is None:
        return []
    return [item.strip() for item in re.split(r"[,;\n]+", str(value)) if item.strip()]


def _join_names(values) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        name = str(value).strip()
        key = name.casefold()
        if name and key not in seen:
            result.append(name)
            seen.add(key)
    return ", ".join(result)


def _parse_date(value: Any) -> date | None:
    try:
        return datetime.strptime(str(value), "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


def _parse_iso_date(value: Any) -> date | None:
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def _parse_clock(value: Any):
    try:
        return datetime.strptime(str(value), "%H:%M").time()
    except (TypeError, ValueError):
        return datetime.strptime("00:00", "%H:%M").time()


def _shift_interval(day_value: date, code: str, settings: dict):
    prefix = {"F": "f", "S": "s", "N": "n"}[code]
    start = datetime.combine(day_value, _parse_clock(settings.get(f"{prefix}_start", "00:00")))
    end = datetime.combine(day_value, _parse_clock(settings.get(f"{prefix}_ende", "00:00")))
    if end <= start:
        end += timedelta(days=1)
    return start, end


def _target_hours(person: dict) -> float:
    try:
        return max(0.0, float(person.get("monatssollstunden", 0.0)))
    except (TypeError, ValueError):
        return 0.0


def _allowed_codes(person: dict) -> set[str]:
    label = str(person.get("feste_dienstart", "Alle Dienste")).strip()
    return set(ALLOWED_SHIFT_CODES.get(label, SHIFT_CODES))


def _rule_enabled(settings: dict, key: str, default: bool = True) -> bool:
    return bool(settings.get(key, default))


def _int_value(value: Any, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return int(default)


def _active_custom_rules(settings: dict) -> list[dict]:
    values = settings.get("eigene_planungsregeln", [])
    if not isinstance(values, list):
        return []
    return [
        dict(rule)
        for rule in values
        if isinstance(rule, dict) and bool(rule.get("aktiv", True))
    ]


def _rule_applies_to(rule: dict, person: dict) -> bool:
    person_id = _int_value(rule.get("mitarbeitende_id", 0), 0)
    return person_id == 0 or person_id == _int_value(person.get("id", -1), -1)


def _rule_weekdays(rule: dict) -> set[int]:
    values = rule.get("wochentage", [])
    if not isinstance(values, list):
        return set()
    result = set()
    for value in values:
        weekday = _int_value(value, -1)
        if 0 <= weekday <= 6:
            result.add(weekday)
    return result


def _weekend_anchor(day_value: date) -> date | None:
    if day_value.weekday() == 5:
        return day_value
    if day_value.weekday() == 6:
        return day_value - timedelta(days=1)
    return None


def _status_name(cp_model, status: int) -> str:
    names = {
        cp_model.OPTIMAL: "OPTIMAL",
        cp_model.FEASIBLE: "FEASIBLE",
        cp_model.INFEASIBLE: "INFEASIBLE",
        cp_model.MODEL_INVALID: "MODEL_INVALID",
        cp_model.UNKNOWN: "UNKNOWN",
    }
    return names.get(status, "UNKNOWN")


def blocked_absence_dates(vorgaben) -> set[tuple[int, date]]:
    """Liefert vollständige Abwesenheitszeiträume inklusive Wochenenden."""
    blocked: set[tuple[int, date]] = set()
    for absence in vorgaben or []:
        if not isinstance(absence, dict):
            continue
        if str(absence.get("code", "")).strip() not in {"U", "W", "FB", "K"}:
            continue
        employee_id = _int_value(absence.get("mitarbeitende_id", -1), -1)
        start = _parse_iso_date(absence.get("startdatum"))
        end = _parse_iso_date(absence.get("enddatum"))
        if employee_id < 0 or start is None or end is None or start > end:
            continue
        current = start
        while current <= end:
            blocked.add((employee_id, current))
            current += timedelta(days=1)
    return blocked


def _blocked_days_from_absences(vorgaben, people, days) -> set[tuple[int, int]]:
    """Übersetzt Abwesenheitsdaten in die Indizes des aktuellen Monatsmodells."""
    employee_by_id = {
        _int_value(person.get("id", -1), -1): employee_index
        for employee_index, person in enumerate(people)
    }
    day_index_by_date = {
        day_value: day_index
        for day_index, day_value in enumerate(days)
        if day_value is not None
    }
    blocked: set[tuple[int, int]] = set()
    for employee_id, day_value in blocked_absence_dates(vorgaben):
        employee_index = employee_by_id.get(employee_id)
        day_index = day_index_by_date.get(day_value)
        if employee_index is None or day_index is None:
            continue
        blocked.add((employee_index, day_index))
    return blocked


def solve_month_v2(plan, employees, settings, vorgaben=None) -> tuple[list[dict], dict]:
    """Plant einen kompletten Monat global mit dem CP-SAT-Solver.

    Harte Regeln werden immer eingehalten. Reicht das Personal dafuer nicht aus,
    verwendet das Modell Strafvariablen fuer offene Besetzungen. So erhaelt die
    Benutzerin trotzdem den bestmoeglichen, pruefbaren Plan.
    """

    cp_model = _cp_model_module()
    prepared_plan = [dict(row) for row in deepcopy(plan or [])]
    people = [dict(person) for person in employees or []]
    if not people:
        return prepared_plan, {
            "engine": "OR-Tools CP-SAT",
            "status": "NO_EMPLOYEES",
            "assigned": 0,
            "fixed": 0,
            "shortages": ["Keine Mitarbeitenden vorhanden."],
        }

    days = [_parse_date(row.get("Datum")) for row in prepared_plan]
    if not prepared_plan or any(day_value is None for day_value in days):
        return prepared_plan, {
            "engine": "OR-Tools CP-SAT",
            "status": "INVALID_INPUT",
            "assigned": 0,
            "fixed": 0,
            "shortages": ["Der Monatsplan enthält ein ungültiges Datum."],
        }

    name_to_employee: dict[str, int] = {}
    duplicate_names = []
    for employee_index, person in enumerate(people):
        key = str(person.get("name", "")).strip().casefold()
        if not key or key in name_to_employee:
            duplicate_names.append(str(person.get("name", "Unbekannt")))
        else:
            name_to_employee[key] = employee_index
    if duplicate_names:
        return prepared_plan, {
            "engine": "OR-Tools CP-SAT",
            "status": "INVALID_INPUT",
            "assigned": 0,
            "fixed": 0,
            "shortages": ["Mitarbeitendennamen müssen eindeutig sein."],
        }

    model = cp_model.CpModel()
    employee_range = range(len(people))
    day_range = range(len(prepared_plan))
    shift_range = range(len(SHIFT_CODES))
    shift_index = {code: index for index, code in enumerate(SHIFT_CODES)}
    x = {
        (employee_index, day_index, code_index): model.NewBoolVar(
            f"x_e{employee_index}_d{day_index}_s{code_index}"
        )
        for employee_index in employee_range
        for day_index in day_range
        for code_index in shift_range
    }

    unavailable: set[tuple[int, int]] = set()
    fixed_codes: dict[tuple[int, int], str] = {}
    unknown_names: list[str] = []
    paid_absence_days = {employee_index: 0 for employee_index in employee_range}

    for day_index, row in enumerate(prepared_plan):
        for column in ABSENCE_COLUMNS:
            for name in _split_names(row.get(column, "")):
                employee_index = name_to_employee.get(name.casefold())
                if employee_index is None:
                    unknown_names.append(name)
                    continue
                unavailable.add((employee_index, day_index))
                if column in PAID_ABSENCE_COLUMNS and days[day_index].weekday() < 5:
                    paid_absence_days[employee_index] += 1
        for code, column in CODE_TO_COLUMN.items():
            for name in _split_names(row.get(column, "")):
                employee_index = name_to_employee.get(name.casefold())
                if employee_index is None:
                    unknown_names.append(name)
                    continue
                current = fixed_codes.get((employee_index, day_index))
                if current and current != code:
                    return prepared_plan, {
                        "engine": "OR-Tools CP-SAT",
                        "status": "INVALID_INPUT",
                        "assigned": 0,
                        "fixed": 0,
                        "shortages": [
                            f'{people[employee_index]["name"]} hat mehrere Dienste am '
                            f'{days[day_index].strftime("%d.%m.%Y")}.'
                        ],
                    }
                fixed_codes[(employee_index, day_index)] = code

    # Im sichtbaren Plan bleibt Urlaub am Wochenende leer. Die ursprüngliche
    # Vorgabe sperrt dennoch den vollständigen Zeitraum für die Planung.
    unavailable.update(_blocked_days_from_absences(vorgaben, people, days))

    targets = [_target_hours(person) for person in people]
    allowed = [_allowed_codes(person) for person in people]
    is_fachkraft = [
        str(person.get("qualifikation", "")) in FACHKRAFT_QUALIFIKATIONEN
        for person in people
    ]
    hours_units = {
        "F": max(0, round(float(settings.get("f_stunden", 0.0)) * HOUR_SCALE)),
        "S": max(0, round(float(settings.get("s_stunden", 0.0)) * HOUR_SCALE)),
        "N": max(0, round(float(settings.get("n_stunden", 0.0)) * HOUR_SCALE)),
    }
    demand = {
        "F": max(0, _int_value(settings.get("min_f", 0), 0)),
        "S": max(0, _int_value(settings.get("min_s", 0), 0)),
        "N": max(0, _int_value(settings.get("min_n", 0), 0)),
    }
    fach_demand = {
        "F": max(0, _int_value(settings.get("min_fach_f", 0), 0)),
        "S": max(0, _int_value(settings.get("min_fach_s", 0), 0)),
        "N": max(0, _int_value(settings.get("min_fach_n", 0), 0)),
    }

    for (employee_index, day_index), code in fixed_codes.items():
        person = people[employee_index]
        invalid_reason = None
        if (employee_index, day_index) in unavailable:
            invalid_reason = "ist an diesem Tag abwesend"
        elif targets[employee_index] <= 0:
            invalid_reason = "hat keine verfügbaren Sollstunden"
        elif code not in allowed[employee_index]:
            invalid_reason = "darf diese Dienstart nicht übernehmen"
        elif code == "N" and str(person.get("nachtdienst", "Nein")).strip() != "Ja":
            invalid_reason = "darf keinen Nachtdienst übernehmen"
        if invalid_reason:
            return prepared_plan, {
                "engine": "OR-Tools CP-SAT",
                "status": "INVALID_INPUT",
                "assigned": 0,
                "fixed": len(fixed_codes),
                "shortages": [
                    f'{person.get("name", "Unbekannt")} {invalid_reason} '
                    f'({days[day_index].strftime("%d.%m.%Y")}).'
                ],
            }

    for employee_index, person in enumerate(people):
        for day_index in day_range:
            model.Add(
                sum(x[employee_index, day_index, code_index] for code_index in shift_range)
                <= 1
            )
            fixed_code = fixed_codes.get((employee_index, day_index))
            for code_index, code in enumerate(SHIFT_CODES):
                variable = x[employee_index, day_index, code_index]
                invalid = (
                    targets[employee_index] <= 0
                    or hours_units[code] <= 0
                    or (employee_index, day_index) in unavailable
                    or code not in allowed[employee_index]
                    or (
                        code == "N"
                        and str(person.get("nachtdienst", "Nein")).strip() != "Ja"
                    )
                )
                if invalid:
                    model.Add(variable == 0)
                elif fixed_code:
                    model.Add(variable == int(code == fixed_code))

    night_block_terms = []
    night_preference_terms = []
    weekend_terms = []
    sequence_terms = []
    fairness_terms = []
    shortage_variables: dict[tuple[int, str], Any] = {}
    fach_shortage_variables: dict[tuple[int, str], Any] = {}
    slot_targets: dict[tuple[int, str], int] = {}
    for day_index in day_range:
        for code_index, code in enumerate(SHIFT_CODES):
            fixed_count = sum(
                1
                for employee_index in employee_range
                if fixed_codes.get((employee_index, day_index)) == code
            )
            slot_target = max(demand[code], fixed_count)
            slot_targets[day_index, code] = slot_target
            shortage = model.NewIntVar(0, slot_target, f"short_d{day_index}_{code}")
            # Die Mindestbesetzung ist eine Untergrenze und keine exakte
            # Sollstaerke. Zusaetzliche Mitarbeitende muessen eingeplant werden
            # duerfen, damit das persoenliche Monatssoll erreicht werden kann.
            model.Add(
                sum(x[employee_index, day_index, code_index] for employee_index in employee_range)
                + shortage
                >= slot_target
            )
            shortage_variables[day_index, code] = shortage

            if _rule_enabled(settings, "regel_fachkraft_aktiv"):
                required_fach = min(slot_target, fach_demand[code])
                fach_shortage = model.NewIntVar(
                    0, required_fach, f"fach_short_d{day_index}_{code}"
                )
                model.Add(
                    sum(
                        x[employee_index, day_index, code_index]
                        for employee_index in employee_range
                        if is_fachkraft[employee_index]
                    )
                    + fach_shortage
                    >= required_fach
                )
                fach_shortage_variables[day_index, code] = fach_shortage

    workdays = max(
        1,
        sum(
            1
            for day_number in range(1, calendar.monthrange(days[0].year, days[0].month)[1] + 1)
            if date(days[0].year, days[0].month, day_number).weekday() < 5
        ),
    )
    service_hours_by_employee: dict[int, Any] = {}
    remaining_target_units: dict[int, int] = {}
    target_gap_terms = []
    target_rule_active = _rule_enabled(settings, "regel_sollstunden_aktiv")
    for employee_index in employee_range:
        target_units = round(targets[employee_index] * HOUR_SCALE)
        absence_credit = round(
            paid_absence_days[employee_index]
            * targets[employee_index]
            / workdays
            * HOUR_SCALE
        )
        service_hours = sum(
            x[employee_index, day_index, code_index] * hours_units[code]
            for day_index in day_range
            for code_index, code in enumerate(SHIFT_CODES)
        )
        service_hours_by_employee[employee_index] = service_hours
        remaining_target_units[employee_index] = max(0, target_units - absence_credit)
        remaining_target = remaining_target_units[employee_index]
        if target_rule_active:
            model.Add(service_hours <= remaining_target)
            target_gap = model.NewIntVar(
                0, remaining_target, f"target_gap_e{employee_index}"
            )
            model.Add(target_gap == remaining_target - service_hours)
        else:
            maximum_service = len(days) * max(hours_units.values(), default=0)
            target_gap = model.NewIntVar(
                0,
                max(maximum_service, remaining_target),
                f"target_gap_e{employee_index}",
            )
            model.AddAbsEquality(target_gap, service_hours - remaining_target)
        target_gap_terms.append(target_gap)

    if _rule_enabled(settings, "regel_ruhezeit_aktiv"):
        min_rest = max(0.0, float(settings.get("ruhezeit", 11.0)))
        for employee_index in employee_range:
            for day_index in range(len(days) - 1):
                for first_code_index, first_code in enumerate(SHIFT_CODES):
                    _, first_end = _shift_interval(days[day_index], first_code, settings)
                    for second_code_index, second_code in enumerate(SHIFT_CODES):
                        second_start, _ = _shift_interval(days[day_index + 1], second_code, settings)
                        rest = (second_start - first_end).total_seconds() / 3600.0
                        if rest < min_rest:
                            model.Add(
                                x[employee_index, day_index, first_code_index]
                                + x[employee_index, day_index + 1, second_code_index]
                                <= 1
                            )

    work_variables = {
        (employee_index, day_index): sum(
            x[employee_index, day_index, code_index] for code_index in shift_range
        )
        for employee_index in employee_range
        for day_index in day_range
    }

    block_start_variables: dict[tuple[int, int, int], Any] = {}
    block_end_variables: dict[tuple[int, int, int], Any] = {}

    def block_start_variable(employee_index: int, day_index: int, code_index: int):
        key = (employee_index, day_index, code_index)
        if key in block_start_variables:
            return block_start_variables[key]
        variable = model.NewBoolVar(
            f"block_start_e{employee_index}_d{day_index}_s{code_index}"
        )
        current = x[employee_index, day_index, code_index]
        if day_index == 0:
            model.Add(variable == current)
        else:
            previous = x[employee_index, day_index - 1, code_index]
            model.Add(variable >= current - previous)
            model.Add(variable <= current)
            model.Add(variable <= 1 - previous)
        block_start_variables[key] = variable
        return variable

    def block_end_variable(employee_index: int, day_index: int, code_index: int):
        key = (employee_index, day_index, code_index)
        if key in block_end_variables:
            return block_end_variables[key]
        variable = model.NewBoolVar(
            f"block_end_e{employee_index}_d{day_index}_s{code_index}"
        )
        current = x[employee_index, day_index, code_index]
        if day_index == len(days) - 1:
            model.Add(variable == current)
        else:
            following = x[employee_index, day_index + 1, code_index]
            model.Add(variable >= current - following)
            model.Add(variable <= current)
            model.Add(variable <= 1 - following)
        block_end_variables[key] = variable
        return variable
    if _rule_enabled(settings, "regel_max_arbeitstage_aktiv"):
        maximum = max(1, _int_value(settings.get("max_arbeitstage", 6), 6))
        for employee_index in employee_range:
            for start in range(max(0, len(days) - maximum)):
                model.Add(
                    sum(
                        work_variables[employee_index, day_index]
                        for day_index in range(start, start + maximum + 1)
                    )
                    <= maximum
                )

    if _rule_enabled(settings, "regel_max_nachtdienste_aktiv"):
        maximum = max(1, _int_value(settings.get("max_nachtdienste", 4), 4))
        night_index = shift_index["N"]
        for employee_index in employee_range:
            for start in range(max(0, len(days) - maximum)):
                model.Add(
                    sum(
                        x[employee_index, day_index, night_index]
                        for day_index in range(start, start + maximum + 1)
                    )
                    <= maximum
                )

    weekend_days: dict[date, list[int]] = {}
    for day_index, day_value in enumerate(days):
        anchor = _weekend_anchor(day_value)
        if anchor is not None:
            weekend_days.setdefault(anchor, []).append(day_index)
    weekend_anchors = sorted(weekend_days)
    weekend_work: dict[tuple[int, date], Any] = {}
    for employee_index in employee_range:
        for anchor in weekend_anchors:
            weekend_variable = model.NewBoolVar(f"weekend_e{employee_index}_{anchor.isoformat()}")
            weekend_work[employee_index, anchor] = weekend_variable
            variables = [
                x[employee_index, day_index, code_index]
                for day_index in weekend_days[anchor]
                for code_index in shift_range
            ]
            for variable in variables:
                model.Add(weekend_variable >= variable)
            model.Add(weekend_variable <= sum(variables))

        if _rule_enabled(settings, "regel_max_wochenenden_aktiv"):
            maximum = max(0, _int_value(settings.get("max_wochenenden", 2), 2))
            model.Add(
                sum(weekend_work[employee_index, anchor] for anchor in weekend_anchors)
                <= maximum
            )
        if bool(settings.get("keine_folgewochenenden", True)):
            for previous, current in zip(weekend_anchors, weekend_anchors[1:]):
                if current == previous + timedelta(days=7):
                    model.Add(
                        weekend_work[employee_index, previous]
                        + weekend_work[employee_index, current]
                        <= 1
                    )

    active_people = [index for index, target in enumerate(targets) if target > 0]
    if active_people and weekend_anchors and _rule_enabled(settings, "regel_wochenende_gleich_aktiv"):
        maximum_weekends = len(weekend_anchors)
        weekend_counts = []
        for employee_index in active_people:
            count = model.NewIntVar(0, maximum_weekends, f"weekend_count_e{employee_index}")
            model.Add(
                count
                == sum(
                    weekend_work[employee_index, anchor]
                    for anchor in weekend_anchors
                )
            )
            weekend_counts.append(count)
        max_count = model.NewIntVar(0, maximum_weekends, "max_weekend_count")
        min_count = model.NewIntVar(0, maximum_weekends, "min_weekend_count")
        model.AddMaxEquality(max_count, weekend_counts)
        model.AddMinEquality(min_count, weekend_counts)
        weekend_terms.append((max_count - min_count) * 10)

    employee_by_id = {
        _int_value(person.get("id", -1), -1): employee_index
        for employee_index, person in enumerate(people)
    }
    for rule in _active_custom_rules(settings):
        rule_type = str(rule.get("typ", ""))
        if rule_type == "nicht_gemeinsam":
            first = employee_by_id.get(_int_value(rule.get("mitarbeitende_a_id", -1), -1))
            second = employee_by_id.get(_int_value(rule.get("mitarbeitende_b_id", -1), -1))
            requested_code = str(rule.get("dienst", "ALL"))
            codes = [requested_code] if requested_code in SHIFT_CODES else list(SHIFT_CODES)
            if first is None or second is None:
                continue
            for day_index in day_range:
                for code in codes:
                    code_index = shift_index[code]
                    model.Add(
                        x[first, day_index, code_index]
                        + x[second, day_index, code_index]
                        <= 1
                    )

        elif rule_type == "frei_nach_dienst":
            trigger_code = str(rule.get("dienst", "N"))
            if trigger_code not in SHIFT_CODES:
                continue
            free_days = max(1, _int_value(rule.get("freie_tage", 1), 1))
            trigger_index = shift_index[trigger_code]
            for employee_index, person in enumerate(people):
                if not _rule_applies_to(rule, person):
                    continue
                for day_index in day_range:
                    # Nachtdienste werden als zusammenhängender Block behandelt.
                    # Freie Tage beginnen deshalb erst nach der letzten Nacht und
                    # nicht nach jeder einzelnen Nacht innerhalb des Blocks.
                    trigger = (
                        block_end_variable(employee_index, day_index, trigger_index)
                        if trigger_code == "N"
                        else x[employee_index, day_index, trigger_index]
                    )
                    for offset in range(1, free_days + 1):
                        following_day = day_index + offset
                        if following_day >= len(days):
                            break
                        model.Add(
                            trigger + work_variables[employee_index, following_day]
                            <= 1
                        )

        elif rule_type == "max_dienstfolge":
            target_code = str(rule.get("dienst", "N"))
            if target_code not in SHIFT_CODES:
                continue
            maximum = max(1, _int_value(rule.get("maximum", 1), 1))
            code_index = shift_index[target_code]
            for employee_index, person in enumerate(people):
                if not _rule_applies_to(rule, person):
                    continue
                for start in range(max(0, len(days) - maximum)):
                    model.Add(
                        sum(
                            x[employee_index, day_index, code_index]
                            for day_index in range(start, start + maximum + 1)
                        )
                        <= maximum
                    )

        elif rule_type == "nicht_an_wochentagen":
            blocked_weekdays = _rule_weekdays(rule)
            for employee_index, person in enumerate(people):
                if not _rule_applies_to(rule, person):
                    continue
                for day_index, day_value in enumerate(days):
                    if day_value.weekday() in blocked_weekdays:
                        for code_index in shift_range:
                            model.Add(x[employee_index, day_index, code_index] == 0)

    if _rule_enabled(settings, "regel_nachtblock_aktiv"):
        night_index = shift_index["N"]
        maximum_nights = max(1, _int_value(settings.get("max_nachtdienste", 4), 4))
        minimum_nights = max(1, _int_value(settings.get("min_nachtblock", 2), 2))
        if _rule_enabled(settings, "regel_max_nachtdienste_aktiv"):
            minimum_nights = min(maximum_nights, minimum_nights)
        free_after_block = max(
            0,
            _int_value(settings.get("freie_tage_nach_nachtblock", 2), 2),
        )
        for employee_index in employee_range:
            for day_index in day_range:
                block_start = block_start_variable(employee_index, day_index, night_index)
                block_end = block_end_variable(employee_index, day_index, night_index)
                night_block_terms.append(block_start)

                # Ein begonnener Nachtblock muss die eingestellte Mindestlänge
                # erreichen. Dadurch entstehen keine isolierten Einzelnächte.
                for offset in range(1, minimum_nights):
                    block_day = day_index + offset
                    if block_day >= len(days):
                        model.Add(block_start == 0)
                        break
                    model.Add(x[employee_index, block_day, night_index] >= block_start)

                # Nach dem Ende des gesamten Nachtblocks folgen echte freie Tage.
                for offset in range(1, free_after_block + 1):
                    following_day = day_index + offset
                    if following_day >= len(days):
                        break
                    model.Add(
                        block_end + work_variables[employee_index, following_day] <= 1
                    )

        if _rule_enabled(settings, "regel_nacht_zuerst_aktiv"):
            for employee_index, person in enumerate(people):
                if str(person.get("feste_dienstart", "")).strip() == "Nur Nachtdienst":
                    continue
                for day_index in day_range:
                    # Reine Nachtwachen werden zuerst genutzt; flexible Personen
                    # bleiben für Früh- und Spätdienste verfügbar.
                    night_preference_terms.append(x[employee_index, day_index, night_index])

    if _rule_enabled(settings, "regel_wochenende_gleich_aktiv"):
        for employee_index in employee_range:
            for anchor, indices in weekend_days.items():
                if len(indices) != 2:
                    continue
                saturday, sunday = sorted(indices)
                single_day = model.NewBoolVar(
                    f"weekend_single_e{employee_index}_{anchor.isoformat()}"
                )
                model.Add(
                    single_day
                    >= work_variables[employee_index, saturday]
                    - work_variables[employee_index, sunday]
                )
                model.Add(
                    single_day
                    >= work_variables[employee_index, sunday]
                    - work_variables[employee_index, saturday]
                )
                weekend_terms.append(single_day * 5)
                for code_index in shift_range:
                    mismatch = model.NewBoolVar(
                        f"weekend_mismatch_e{employee_index}_{anchor.isoformat()}_s{code_index}"
                    )
                    model.Add(
                        mismatch
                        >= x[employee_index, saturday, code_index]
                        - x[employee_index, sunday, code_index]
                    )
                    model.Add(
                        mismatch
                        >= x[employee_index, sunday, code_index]
                        - x[employee_index, saturday, code_index]
                    )
                    weekend_terms.append(mismatch * 2)

    # Unruhige direkte Wechsel zwischen Früh- und Spätdienst werden vermieden.
    # Es bleibt eine weiche Präferenz, damit eine notwendige Besetzung möglich ist.
    early_index = shift_index["F"]
    late_index = shift_index["S"]
    for employee_index in employee_range:
        for day_index in range(len(days) - 1):
            transition = model.NewBoolVar(
                f"fs_transition_e{employee_index}_d{day_index}"
            )
            model.Add(
                transition
                >= x[employee_index, day_index, early_index]
                + x[employee_index, day_index + 1, late_index]
                - 1
            )
            model.Add(
                transition
                >= x[employee_index, day_index, late_index]
                + x[employee_index, day_index + 1, early_index]
                - 1
            )
            sequence_terms.append(transition)

    for employee_index, person in enumerate(people):
        if len(allowed[employee_index]) <= 1:
            continue
        early_count = sum(
            x[employee_index, day_index, shift_index["F"]]
            for day_index in day_range
        )
        late_count = sum(
            x[employee_index, day_index, shift_index["S"]]
            for day_index in day_range
        )
        imbalance = model.NewIntVar(0, len(days), f"fs_imbalance_e{employee_index}")
        model.AddAbsEquality(imbalance, early_count - late_count)
        fairness_terms.append(imbalance)

    # Verbleibende unvermeidbare Minusstunden werden proportional verteilt. Das
    # persoenliche Rest-Soll selbst bleibt das Ziel; die Mindestbesetzung darf
    # den Solver nicht mehr schon vorher stoppen.
    required_service_units = sum(
        slot_targets[day_index, code] * hours_units[code]
        for day_index in day_range
        for code in SHIFT_CODES
    )
    total_remaining_units = sum(remaining_target_units.values())
    relative_deviations = []
    absolute_deviations = []
    relative_upper_bound = 10_000
    if total_remaining_units > 0:
        maximum_hours = max(required_service_units, total_remaining_units, 1)
        smallest_remaining = min(
            value for value in remaining_target_units.values() if value > 0
        )
        relative_upper_bound = max(
            10_000,
            (maximum_hours * 1_000 + smallest_remaining - 1) // smallest_remaining,
        )
        for employee_index in employee_range:
            remaining = remaining_target_units[employee_index]
            if remaining <= 0:
                continue
            desired = remaining
            deviation = model.NewIntVar(
                0, maximum_hours, f"fair_abs_e{employee_index}"
            )
            model.AddAbsEquality(
                deviation,
                service_hours_by_employee[employee_index] - desired,
            )
            relative = model.NewIntVar(
                0, relative_upper_bound, f"fair_rel_e{employee_index}"
            )
            model.Add(relative * remaining >= deviation * 1_000)
            relative_deviations.append(relative)
            absolute_deviations.append(deviation)

    if relative_deviations:
        max_relative = model.NewIntVar(0, relative_upper_bound, "fair_max_relative")
        model.AddMaxEquality(max_relative, relative_deviations)
        absolute_upper = len(absolute_deviations) * maximum_hours
        fs_imbalance_upper = len(fairness_terms) * len(days)
        lower_priority_upper = absolute_upper + fs_imbalance_upper
        relative_sum_upper = len(relative_deviations) * relative_upper_bound
        relative_weight = lower_priority_upper + 1
        maximum_weight = (
            relative_sum_upper * relative_weight + lower_priority_upper + 1
        )
        fairness_terms.append(max_relative * maximum_weight)
        fairness_terms.append(sum(relative_deviations) * relative_weight)
    fairness_terms.append(sum(absolute_deviations))

    # Echte Prioritätsstufen wie in professioneller Dienstplanung:
    # 1. Besetzung/Fachkräfte, 2. persoenliches Monatssoll,
    # 3. Nachtblöcke/Wochenenden, 4. faire Verteilung der Restabweichungen.
    total_fach_upper = sum(
        min(slot_targets[day_index, code], fach_demand[code])
        for day_index in day_range
        for code in SHIFT_CODES
    )
    coverage_objective = (
        sum(shortage_variables.values()) * (total_fach_upper + 1)
        + sum(fach_shortage_variables.values())
    )
    target_objective = sum(target_gap_terms)

    weekend_upper = max(1, len(people) * max(1, len(weekend_anchors)) * 20)
    sequence_upper = max(1, len(people) * max(1, len(days) - 1))
    preference_upper = max(1, len(people) * len(days))
    preference_weight = weekend_upper + sequence_upper + 1
    block_weight = preference_upper * preference_weight + weekend_upper + sequence_upper + 1
    structure_objective = (
        sum(night_block_terms) * block_weight
        + sum(night_preference_terms) * preference_weight
        + sum(weekend_terms) * (sequence_upper + 1)
        + sum(sequence_terms)
    )
    fairness_objective = sum(fairness_terms)

    max_seconds = max(
        1.0, min(60.0, float(settings.get("solver_v2_max_seconds", 20.0)))
    )
    phases = [
        ("Besetzung", coverage_objective),
        ("Sollstunden", target_objective),
        ("Nachtblöcke und Wochenenden", structure_objective),
        ("Fairness", fairness_objective),
    ]
    phase_seconds = max(0.25, max_seconds / len(phases))
    phase_scores: dict[str, int] = {}
    phase_statuses: dict[str, str] = {}
    solve_seconds = 0.0
    solver = None
    status = cp_model.UNKNOWN

    for phase_index, (phase_name, phase_objective) in enumerate(phases):
        model.Minimize(phase_objective)
        phase_solver = cp_model.CpSolver()
        phase_solver.parameters.max_time_in_seconds = phase_seconds
        # Bewusst deterministisch und ohne zusaetzliche Parallelverarbeitung.
        phase_solver.parameters.num_search_workers = 1
        phase_solver.parameters.random_seed = 2026
        phase_status = phase_solver.Solve(model)
        solve_seconds += phase_solver.WallTime()
        phase_statuses[phase_name] = _status_name(cp_model, phase_status)
        if phase_status not in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
            if solver is None:
                solver = phase_solver
                status = phase_status
            break

        solver = phase_solver
        status = phase_status
        score = int(round(phase_solver.ObjectiveValue()))
        phase_scores[phase_name] = score
        if phase_index < len(phases) - 1:
            # Spätere Phasen dürfen ein besseres Ergebnis finden, aber niemals
            # eine bereits erreichte wichtigere Priorität verschlechtern.
            model.Add(phase_objective <= score)

    status_name = _status_name(cp_model, status)
    if status not in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
        return prepared_plan, {
            "engine": "OR-Tools CP-SAT",
            "status": status_name,
            "assigned": 0,
            "fixed": len(fixed_codes),
            "shortages": [
                "OR-Tools konnte mit den aktiven harten Regeln keinen gültigen Plan berechnen."
            ],
            "solve_seconds": round(solve_seconds, 2),
            "phases": phase_statuses,
        }

    solved_plan = [dict(row) for row in prepared_plan]
    for row in solved_plan:
        for column in CODE_TO_COLUMN.values():
            row[column] = ""

    assigned_total = 0
    fixed_total = 0
    for day_index, row in enumerate(solved_plan):
        for code_index, code in enumerate(SHIFT_CODES):
            names = []
            for employee_index, person in enumerate(people):
                if solver.Value(x[employee_index, day_index, code_index]):
                    names.append(str(person.get("name", "")))
                    if fixed_codes.get((employee_index, day_index)) == code:
                        fixed_total += 1
                    else:
                        assigned_total += 1
            row[CODE_TO_COLUMN[code]] = _join_names(names)
        row["Status"] = "Intelligent mit OR-Tools geplant"

    shortages = []
    for day_index, day_value in enumerate(days):
        for code in SHIFT_CODES:
            missing = solver.Value(shortage_variables[day_index, code])
            if missing:
                shortages.append(
                    f'{day_value.strftime("%d.%m.%Y")}: {code} – '
                    f'{missing} offene Besetzung{"en" if missing != 1 else ""}.'
                )
            fach_variable = fach_shortage_variables.get((day_index, code))
            if fach_variable is not None:
                missing_fach = solver.Value(fach_variable)
                if missing_fach:
                    shortages.append(
                        f'{day_value.strftime("%d.%m.%Y")}: {code} – '
                        f'{missing_fach} Fachkraft-Besetzung{"en" if missing_fach != 1 else ""} offen.'
                    )

    required_hours = len(days) * sum(
        demand[code] * hours_units[code] / HOUR_SCALE for code in SHIFT_CODES
    )
    return solved_plan, {
        "engine": "OR-Tools CP-SAT",
        "status": status_name,
        "assigned": assigned_total,
        "fixed": fixed_total,
        "shortages": shortages,
        "required_hours": round(required_hours, 1),
        "available_hours": round(sum(targets), 1),
        "solve_seconds": round(solve_seconds, 2),
        "objective": phase_scores,
        "phases": phase_statuses,
        "unknown_names": sorted(set(unknown_names), key=str.casefold),
    }
