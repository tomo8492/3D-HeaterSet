# schedule_utils.py
"""Pure helper functions for reschedule request feature. No NiceGUI/Oracle deps."""
from datetime import date, timedelta


def add_workdays(start_date: date, n: int) -> date:
    """
    Return the date that is n workdays after start_date.
    Skips Saturday (weekday 5) and Sunday (weekday 6).
    n=0 returns start_date unchanged.

    # TODO: query FactoryCalendar here when internal server calendar is identified.
    # Current behaviour: weekends-only fallback (Japanese national holidays NOT excluded).
    """
    if n == 0:
        return start_date
    current = start_date
    added = 0
    while added < n:
        current += timedelta(days=1)
        if current.weekday() < 5:  # Mon-Fri
            added += 1
    return current


def _group_by_date(rows: list) -> list[dict]:
    """
    From table rows, produce groups of consecutive modified rows sharing the same date.

    Args:
        rows: list of row dicts with keys '#', '現在工程', '_color', '_raw_date'.
              Only rows where _color != '' are included.

    Returns:
        List of {'name': str, 'date': str} dicts.
        name is 'ProcessA' for a single row, 'ProcessA~ProcessZ' for a range.
        date is YYYY-MM-DD (raw HTML date format).
    """
    modified = sorted(
        [r for r in rows if r.get('_color', '') != ''],
        key=lambda r: r['#']
    )
    groups = []
    i = 0
    while i < len(modified):
        d = modified[i]['_raw_date']
        j = i
        while j + 1 < len(modified) and modified[j + 1]['_raw_date'] == d:
            j += 1
        if i == j:
            name = modified[i]['現在工程']
        else:
            name = f"{modified[i]['現在工程']}~{modified[j]['現在工程']}"
        groups.append({'name': name, 'date': d})
        i = j + 1
    return groups


def _build_request_text(rows: list) -> str:
    """
    Generate formatted reschedule request string from table rows.

    Args:
        rows: list of row dicts (same format as _group_by_date).

    Returns:
        Multi-line string ending with 'でリスケお願いします。'.
        Returns 'でリスケお願いします。' alone if no modified rows exist.
    """
    groups = _group_by_date(rows)
    if not groups:
        return 'でリスケお願いします。'
    lines = [f"{g['name']} : {g['date'].replace('-', '/')}" for g in groups]
    lines.append('でリスケお願いします。')
    return '\n'.join(lines)


def _build_request_text_from_list(dropdown_rows: list[dict]) -> str:
    """
    Generate reschedule request string from dropdown state list.

    Consecutive rows sharing the same date are merged: 'A~B : date'.
    Rows with value='' are skipped (not set / deleted).

    Args:
        dropdown_rows: list of {'name': str, 'value': str} dicts, in order.

    Returns:
        Multi-line string ending with 'でリスケお願いします。',
        or '' if no rows have a date set (textarea stays blank).
    """
    active = [r for r in dropdown_rows if r.get('value', '') and r.get('checked', True)]
    if not active:
        return ''
    groups = []
    i = 0
    while i < len(active):
        d = active[i]['value']
        j = i
        while j + 1 < len(active) and active[j + 1]['value'] == d:
            j += 1
        name = active[i]['name'] if i == j else f"{active[i]['name']}~{active[j]['name']}"
        groups.append(f"{name} : {d.replace('-', '/')}")
        i = j + 1
    groups.append('でリスケお願いします。')
    return '\n'.join(groups)


def cascade_dates_with_capacity(
    trigger_name: str,
    trigger_date: date,
    subsequent_names: list[str],
) -> list[date]:
    """
    Assign workday dates to a sequence of subsequent processes given that
    `trigger_name` is already scheduled on `trigger_date`.

    Capacity rules (per process type, per day — types are independent):
      - Names containing '加工': max 2 per day  (heavier machining step)
      - All other names:        max 3 per day

    The trigger process itself counts as one unit toward its type-bucket on
    `trigger_date`. Dates are monotonically non-decreasing.
    """
    def _type_cap(name: str) -> tuple[str, int]:
        if '加工' in name:
            return '加工', 2
        return 'other', 3

    # {(iso_date_str, type_key): usage_count}
    usage: dict[tuple[str, str], int] = {}

    # Normalise trigger_date to the nearest following workday
    while trigger_date.weekday() >= 5:
        trigger_date += timedelta(days=1)

    # Seed trigger process into usage
    t_type, _ = _type_cap(trigger_name)
    usage[(trigger_date.isoformat(), t_type)] = 1

    cur_day = trigger_date
    result: list[date] = []

    for name in subsequent_names:
        type_key, cap = _type_cap(name)

        # Find first workday >= cur_day with capacity for this type
        try_day = cur_day
        while True:
            day_key = (try_day.isoformat(), type_key)
            if usage.get(day_key, 0) < cap:
                break
            try_day = add_workdays(try_day, 1)

        day_key = (try_day.isoformat(), type_key)
        usage[day_key] = usage.get(day_key, 0) + 1
        result.append(try_day)
        cur_day = try_day  # subsequent processes cannot go before this day

    return result
