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


def solve_month_v2(plan, employees, settings) -> tuple[list[dict], dict]:
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

    objective_terms = []
    shortage_variables: dict[tuple[int, str], Any] = {}
    fach_shortage_variables: dict[tuple[int, str], Any] = {}
    for day_index in day_range:
        for code_index, code in enumerate(SHIFT_CODES):
            fixed_count = sum(
                1
                for employee_index in employee_range
                if fixed_codes.get((employee_index, day_index)) == code
            )
            slot_target = max(demand[code], fixed_count)
            shortage = model.NewIntVar(0, slot_target, f"short_d{day_index}_{code}")
            model.Add(
                sum(x[employee_index, day_index, code_index] for employee_index in employee_range)
                + shortage
                == slot_target
            )
            shortage_variables[day_index, code] = shortage
            objective_terms.append(shortage * 100_000)

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
                objective_terms.append(fach_shortage * 80_000)

    workdays = max(
        1,
        sum(
            1
            for day_number in range(1, calendar.monthrange(days[0].year, days[0].month)[1] + 1)
            if date(days[0].year, days[0].month, day_number).weekday() < 5
        ),
    )
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
        if _rule_enabled(settings, "regel_sollstunden_aktiv"):
            model.Add(service_hours <= max(0, target_units - absence_credit))

        max_deviation = max(
            target_units + absence_credit,
            sum(demand[code] * hours_units[code] for code in SHIFT_CODES) * len(days),
            1,
        )
        under = model.NewIntVar(0, max_deviation, f"under_e{employee_index}")
        over = model.NewIntVar(0, max_deviation, f"over_e{employee_index}")
        model.Add(service_hours + absence_credit + under - over == target_units)
        objective_terms.extend((under * 5, over * 25))

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
        objective_terms.append((max_count - min_count) * 200)

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
                    for offset in range(1, free_days + 1):
                        following_day = day_index + offset
                        if following_day >= len(days):
                            break
                        model.Add(
                            x[employee_index, day_index, trigger_index]
                            + work_variables[employee_index, following_day]
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
        for employee_index in employee_range:
            for day_index in day_range:
                block_start = model.NewBoolVar(f"night_start_e{employee_index}_d{day_index}")
                current = x[employee_index, day_index, night_index]
                if day_index == 0:
                    model.Add(block_start == current)
                else:
                    previous = x[employee_index, day_index - 1, night_index]
                    model.Add(block_start >= current - previous)
                    model.Add(block_start <= current)
                    model.Add(block_start <= 1 - previous)
                objective_terms.append(block_start * 30)

    if _rule_enabled(settings, "regel_wochenende_gleich_aktiv"):
        for employee_index in employee_range:
            for anchor, indices in weekend_days.items():
                if len(indices) != 2:
                    continue
                saturday, sunday = sorted(indices)
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
                    objective_terms.append(mismatch * 20)

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
        objective_terms.append(imbalance * 2)

    model.Minimize(sum(objective_terms))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(
        1.0, min(60.0, float(settings.get("solver_v2_max_seconds", 20.0)))
    )
    # Bewusst deterministisch und ohne zusaetzliche Parallelverarbeitung.
    solver.parameters.num_search_workers = 1
    solver.parameters.random_seed = 2026
    status = solver.Solve(model)
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
            "solve_seconds": round(solver.WallTime(), 2),
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
        "solve_seconds": round(solver.WallTime(), 2),
        "objective": round(solver.ObjectiveValue(), 1),
        "unknown_names": sorted(set(unknown_names), key=str.casefold),
    }
