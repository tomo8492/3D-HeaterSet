"""gantt_dialog.py
Shared gantt helpers for the main app batch reschedule dialog.

Exports:
  Pure helpers  : _oracle_df_to_schedule, _trim_to_next_ukeire,
                  _build_flows_and_schedule,
                  _d_to_js, _compute_flow_processes, _build_js_vars
  Constants     : _GANTT_CSS, _GANTT_BODY, _GANTT_JS
  Entry point   : open_gantt_dialog()
"""
import asyncio
import json
import re
from pathlib import Path
from datetime import date, timedelta  # timedelta used by open_gantt_dialog

# ---------------------------------------------------------------------------
# Module constants
# ---------------------------------------------------------------------------
WEEKDAY_JP = ['月', '火', '水', '木', '金', '土', '日']

_FLOW_COLORS = [
    '#4a90d9', '#e67e22', '#27ae60', '#9b59b6',
    '#e74c3c', '#1abc9c', '#f39c12', '#2980b9',
]

# ---------------------------------------------------------------------------
# Pure helpers
# ---------------------------------------------------------------------------

def _oracle_df_to_schedule(lot_id: str, df) -> list:
    """Map one Oracle mode-v DataFrame to [(proc_name, date), ...].

    Args:
        lot_id: Oracle lot ID string (unused in computation; kept for API symmetry).
        df:     DataFrame with columns '現在工程' (process name) and '開始時間'
                (planned start datetime). Rows in Oracle sequence order.
                df.columns must already be lowercased by caller — Japanese chars
                are unaffected by .str.lower().

    Returns:
        [(proc_name, start_date), ...] in row order.
    """
    result = []
    for _, row in df.iterrows():
        proc_name = row['現在工程']
        ts = row['開始時間']
        d = ts.date() if hasattr(ts, 'date') else ts
        result.append((proc_name, d))
    return result


def _trim_to_next_ukeire(entries: list) -> list:
    """Trim schedule entries to the range [current_process … next 受入_ process].

    Rules:
    - Search from index 1 onward for the first process name starting with '受入_'.
    - When found, return entries[:i+1] (current process through that entry, inclusive).
    - If the current process (index 0) itself starts with '受入_', this naturally finds
      the *next* 受入_ after it — the rule is identical regardless.
    - If no qualifying '受入_' is found after index 0, return all entries unchanged.

    Args:
        entries: [(proc_name, date), ...] in Oracle row order.

    Returns:
        Trimmed list; the original list is never mutated.
    """
    if not entries:
        return entries
    for i in range(1, len(entries)):
        if entries[i][0].startswith('受入_'):
            return entries[: i + 1]
    return entries


def _build_flows_and_schedule(groups: dict) -> tuple:
    """Build FLOWS list and flat schedule dict from grouped Oracle results.

    Args:
        groups: {flowname: [(lot_id, productcode, df), ...]}
                productcode is carried through for the Excel write caller;
                it is not used inside this function.

    Returns:
        (flows, schedule) where
          flows    = [{id, label, color, lots:[{id, label}]}, ...]
          schedule = {lot_id: [(proc_name, date), ...]}
    """
    flows = []
    schedule = {}
    for i, (flowname, lot_tuples) in enumerate(groups.items()):
        color = _FLOW_COLORS[i % len(_FLOW_COLORS)]
        lots = []
        for lot_id, _productcode, df in lot_tuples:
            lots.append({'id': lot_id, 'label': lot_id})
            schedule[lot_id] = _trim_to_next_ukeire(_oracle_df_to_schedule(lot_id, df))
        flows.append({
            'id':    flowname,
            'label': flowname,
            'color': color,
            'lots':  lots,
        })
    return flows, schedule


def _d_to_js(d: date) -> str:
    """Convert a date to a JS object literal string used in DATES / FIXED_DATES constants."""
    return (
        '{iso:"' + d.isoformat() + '",'
        'label:"' + d.strftime('%m/%d') + '",'
        'dow:' + str(d.weekday()) + ','
        'dowLabel:"' + WEEKDAY_JP[d.weekday()] + '"}'
    )


def _compute_flow_processes(flows: list, schedule: dict) -> dict:
    """Return {flowId: [proc_name, ...]} — union of all processes per flow.

    Outer loop: lots in flows[i]['lots'] array order.
    Inner loop: schedule[lotId] entries in order.
    Process names are deduplicated, preserving first-appearance order.
    """
    result: dict[str, list[str]] = {}
    for flow in flows:
        seen: list[str] = []
        seen_set: set[str] = set()
        for lot in flow['lots']:
            for proc_name, _ in schedule.get(lot['id'], []):
                if proc_name not in seen_set:
                    seen.append(proc_name)
                    seen_set.add(proc_name)
        result[flow['id']] = seen
    return result


def _build_js_vars(flows: list, today: date, all_dates: list,
                   schedule: dict, fixed_dates: list, page_dates: list) -> str:
    """Build JS constants string injected via /*VARS*/ placeholder in _GANTT_JS.

    Signature differences from gantt_demo.py original:
      - `flows` added as first param  (original read module-level FLOWS global)
      - `today` added as second param (original read module-level TODAY global)
    """
    # FLOWS — grouped structure (for DOM build + output grouping)
    flows_parts = []
    for flow in flows:
        lots_js = '[' + ','.join(
            '{id:"' + lot['id'] + '",label:"' + lot['label'] + '"}'
            for lot in flow['lots']
        ) + ']'
        flows_parts.append(
            '{id:"' + flow['id'] + '",'
            'label:"' + flow['label'] + '",'
            'color:"' + flow['color'] + '",'
            'lots:' + lots_js + '}'
        )
    flows_js = '[' + ','.join(flows_parts) + ']'

    # ALL_LOTS — flat array with color + flowId injected
    all_lots_parts = []
    for flow in flows:
        for lot in flow['lots']:
            all_lots_parts.append(
                '{id:"' + lot['id'] + '",'
                'label:"' + lot['label'] + '",'
                'color:"' + flow['color'] + '",'
                'flowId:"' + flow['id'] + '"}'
            )
    all_lots_js = '[' + ','.join(all_lots_parts) + ']'

    dates_js      = '[' + ','.join(_d_to_js(d) for d in all_dates)   + ']'
    fixed_js      = '[' + ','.join(_d_to_js(d) for d in fixed_dates) + ']'
    page_dates_js = '[' + ','.join(
        '[' + ','.join(_d_to_js(d) for d in page) + ']'
        for page in page_dates
    ) + ']'

    sched_parts, order_parts = [], []
    for flow in flows:
        for lot in flow['lots']:
            procs = schedule.get(lot['id'], [])
            items = ','.join(
                '{name:"' + p + '",date:"' + dt.isoformat() + '"}'
                for p, dt in procs
            )
            sched_parts.append('"' + lot['id'] + '":[' + items + ']')
            order_items = ','.join('"' + p + '"' for p, _ in procs)
            order_parts.append('"' + lot['id'] + '":[' + order_items + ']')

    # FLOW_PROCESSES — union of process names per flow, ordered by first appearance
    fp = _compute_flow_processes(flows, schedule)
    fp_parts = []
    for flow in flows:
        procs_js = '[' + ','.join('"' + p + '"' for p in fp[flow['id']]) + ']'
        fp_parts.append('"' + flow['id'] + '":' + procs_js)
    flow_processes_js = '{' + ','.join(fp_parts) + '}'

    return (
        f'const TODAY="{today.isoformat()}";'
        f'const FLOWS={flows_js};'
        f'const ALL_LOTS={all_lots_js};'
        f'const DATES={dates_js};'
        f'const FIXED_DATES={fixed_js};'
        f'const PAGE_DATES={page_dates_js};'
        f'const INIT_SCHEDULE={{' + ','.join(sched_parts) + '};'
        f'const PROC_ORDER={{' + ','.join(order_parts) + '};'
        f'const FLOW_PROCESSES={flow_processes_js};'
        f'const MAX_PER_CELL=5;'
    )


# ---------------------------------------------------------------------------
# HTML / CSS / JS — loaded from gantt_template.html at import time
# ---------------------------------------------------------------------------

def _load_gantt_template() -> tuple[str, str, str]:
    """Read gantt_template.html and extract the three named sections."""
    path = Path(__file__).parent / 'gantt_template.html'
    text = path.read_text(encoding='utf-8')

    def _extract(name: str) -> str:
        m = re.search(
            r'<!-- \[' + re.escape(name) + r'\] -->(.*?)<!-- \[/' + re.escape(name) + r'\] -->',
            text, re.DOTALL,
        )
        if not m:
            raise ValueError(f'Section [{name}] not found in gantt_template.html')
        return m.group(1).strip('\n')

    return _extract('GANTT_CSS'), _extract('GANTT_BODY'), _extract('GANTT_JS')


_GANTT_CSS, _GANTT_BODY, _GANTT_JS = _load_gantt_template()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

async def open_gantt_dialog(
    lot_id_flow_product_triples: list,
    username: str,
    excel_config: dict,
    query_fn,
    spinner_open=None,
    spinner_close=None,
) -> None:
    """Open the interactive Gantt drag-and-drop dialog for batch reschedule.

    Args:
        lot_id_flow_product_triples: [(lot_id, flowname, productcode), ...]
        username:      Current user ID string (e.g. USER_INS.uid from guitest.py).
        excel_config:  Dict with keys 'path' and 'sheet' for the reschedule Excel file.
        query_fn:      Async callable (mode: str, lot_id: str) -> (dict[str, DataFrame], any).
                       Pass guitest.query_data_stdin.
        spinner_open:  Optional callable to show a loading indicator.
        spinner_close: Optional callable to hide a loading indicator.
    """
    from nicegui import ui
    from modules.xw_excel import append_reschedule_row


    # ── 1. Oracle fetch ──────────────────────────────────────────────────────
    if spinner_open:
        spinner_open()
    lot_schedules: dict = {}
    _failed_lots: list = []
    try:
        for lot_id, flowname, productcode in lot_id_flow_product_triples:
            try:
                result, _ = await query_fn('v', lot_id)
                if result:
                    df = next(iter(result.values()))
                    df.columns = df.columns.str.lower()
                    # Four-element tuple: lot_id stored at [0] so step 2 can unpack cleanly
                    lot_schedules[lot_id] = (lot_id, flowname, productcode, df)
            except Exception:
                _failed_lots.append(lot_id)
    finally:
        if spinner_close:
            spinner_close()

    if _failed_lots:
        ui.notify(f'{len(_failed_lots)}件のデータ取得に失敗しました', type='warning')
    if not lot_schedules:
        ui.notify('データが取得できませんでした', type='warning')
        return

    # ── 2. Group by flowname ─────────────────────────────────────────────────
    groups: dict = {}
    for _lid, flowname, productcode, df in lot_schedules.values():
        groups.setdefault(flowname, []).append((_lid, productcode, df))
    flows, schedule = _build_flows_and_schedule(groups)

    # lot_productcodes used by _on_write to look up productcode per lot
    lot_productcodes: dict = {
        lot_id: productcode
        for lot_id, _, productcode, _ in lot_schedules.values()
    }

    # ── 3. Build date window ─────────────────────────────────────────────────
    today = date.today()
    DAYS_BEFORE      = 5
    FIXED_END_OFFSET = 7
    PAGE_DAYS        = 7
    PAGE_COUNT       = 3

    fixed_dates = [
        today - timedelta(days=DAYS_BEFORE) + timedelta(days=i)
        for i in range(DAYS_BEFORE + FIXED_END_OFFSET + 1)
    ]
    paginated = [
        today + timedelta(days=FIXED_END_OFFSET + 1 + i)
        for i in range(PAGE_DAYS * PAGE_COUNT)
    ]
    page_dates = [
        paginated[i * PAGE_DAYS:(i + 1) * PAGE_DAYS]
        for i in range(PAGE_COUNT)
    ]
    all_dates = fixed_dates + paginated

    # ── 4. Build JS vars ─────────────────────────────────────────────────────
    js_vars = _build_js_vars(flows, today, all_dates, schedule, fixed_dates, page_dates)

    # ── 5. Build anywidget ESM ────────────────────────────────────────────────
    # anywidget's render({ model, el }) fires AFTER the element is mounted, so
    # getElementById works immediately — no Vue timing race.
    import anywidget
    _gantt_body_json = json.dumps(_GANTT_BODY)
    _gantt_code = _GANTT_JS.replace('/*VARS*/', js_vars)
    _esm_string = (
        "function render({ model, el }) {\n"
        "  ['#settings-popover','#output-panel'].forEach("
        "    id => { const e = document.querySelector(id); if (e) e.remove(); }"
        "  );\n"
        "  el.style.display = 'block';\n"
        "  el.style.width = '100%';\n"
        "  el.innerHTML = " + _gantt_body_json + ";\n"
        + _gantt_code + "\n"
        "}\n"
        "export default { render };\n"
    )

    class _GanttWidget(anywidget.AnyWidget):

        _esm = _esm_string  # class attribute carries the ESM per-invocation class


    widget = _GanttWidget()

    # ── 6. Open dialog ───────────────────────────────────────────────────────
    dlg = ui.dialog().props('maximized')

    async def _on_close():
        await ui.run_javascript(
            "['#settings-popover','#output-panel']"
            ".forEach(id => { const el = document.querySelector(id); if (el) el.remove(); })"
        )
        dlg.close()

    async def _on_write():
        try:
            strings = await ui.run_javascript('return gatherOutput()')
        except Exception as exc:
            ui.notify(f'データ取得失敗: {exc}', type='negative')
            return

        lots_to_write = [
            (lot_id, productcode, strings.get(lot_id, ''))
            for lot_id, productcode in lot_productcodes.items()
            if strings.get(lot_id, '')
        ]
        if not lots_to_write:
            ui.notify('書き込むデータがありません', type='warning')
            return

        total = len(lots_to_write)

        with ui.dialog().props(
            'persistent backdrop-filter="blur(5px) brightness(60%)"'
        ) as overlay, ui.card().classes('items-center gap-3 p-6 min-w-64'):
            status_icon  = ui.icon('hourglass_empty', size='xl').classes('text-blue-400')
            status_label = ui.label('Excelを開いています...').classes('text-sm')
            exit_btn = (ui.button('閉じる', icon='close')
                          .props('color=negative flat dense')
                          .classes('mt-2'))
            exit_btn.set_visibility(False)
        overlay.open()

        exited = False
        active_task: asyncio.Task | None = None

        def _on_exit() -> None:
            nonlocal exited
            exited = True
            if active_task and not active_task.done():
                active_task.cancel()

        exit_btn.on_click(_on_exit)

        async def _reveal_exit() -> None:
            await asyncio.sleep(30)
            if not exited:
                exit_btn.set_visibility(True)

        reveal_task = asyncio.create_task(_reveal_exit())

        try:
            written = 0
            for i, (lot_id, productcode, text) in enumerate(lots_to_write, 1):
                if exited:
                    break
                status_label.set_text(f'書き込み中... {i}/{total}')
                active_task = asyncio.create_task(append_reschedule_row(
                    request_str=text,
                    username=username,
                    lotid=lot_id,
                    db_figure=productcode,
                    file_path=excel_config['path'],
                    sheet_name=excel_config['sheet'],
                ))
                await active_task
                written += 1

            if not exited:
                status_icon.props('name=check_circle color=positive size=xl')
                status_label.set_text(f'{written}件書き込み完了')
                await asyncio.sleep(1.5)

        except asyncio.CancelledError:
            pass
        except Exception as exc:
            ui.notify(f'書き出し失敗: {exc}', type='negative')
        finally:
            reveal_task.cancel()
            overlay.close()
            overlay.delete()


    with dlg, ui.card().classes('w-full h-full column'):
        with ui.row().classes('items-center w-full'):
            ui.label('リスケ依頼生成（バッチ）').classes('text-lg font-bold flex-1')
            ui.button('Excelに書き込む', on_click=_on_write).props('color=primary')
            ui.button('閉じる', on_click=_on_close).props('flat')
        with ui.scroll_area().classes('w-full flex-1'):
            ui.anywidget(widget)

    # ── 7. Open dialog ───────────────────────────────────────────────────────
    dlg.open()


