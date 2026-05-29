from datetime import datetime, timedelta, date
from typing import Literal, List
from nicegui import ui
import uuid

ProcessStatus = Literal['done', 'current', 'future', 'empty']

class Gantt:
    def __init__(self):
        self.gantt_container = None
        self.lots_data = None
        self.dates = None
        self.current_page = 0
        self.page_size = 20
        self.start_date = None
        self.end_date = None
        self.cell_width = 80
        self.show_all_labels = True
        # Unique class name for JS targeting
        self._uid = f'gantt-{uuid.uuid4().hex[:8]}'
        self._group_counts = {'soon': 0, 'completed': 0, 'on_hold': 0, 'abnormal': 0}

    def create_gantt_from_raw_data(self,
        lot_data: dict,
        start_date: str = None,
        end_date: str = None,
        container = None,
        page_size: int = 20,
        cell_width: int = 80,
        show_all_labels: bool = True,
        **kwargs
    ):
        """
        Create Gantt chart directly from raw lot data with pagination.
        
        Args:
            lot_data: Dict{ lot_id: List[ Tuple(ProcessName, Timestamp, HoldReason, LeadTime) ] }
            start_date: Optional start date (auto-detected if None)
            end_date: Optional end date (auto-detected if None)
            container: Optional NiceGUI container to place the chart in
            page_size: Number of rows per page (default 20)
            cell_width: Width of each date cell in pixels
            show_all_labels: Show labels for done processes
        """
        #print('create_gantt')
        self.lots_data = self.parse_lot_data(lot_data)
        
        # Sort into priority groups for effective information display:
        # Group 0: Completing soon (future end within 14 days) - by completion date asc
        # Group 1: Completing later (future end beyond 14 days) - by completion date asc
        # Group 2: Completed (no future processes) - by last date desc
        # Group 3: On-hold (current process has hold reason) - by last date desc
        # Group 4: Abnormal (single process, not on hold) - by date desc
        
        today = datetime.now().strftime('%Y/%m/%d')
        soon_cutoff = (datetime.now() + timedelta(days=14)).strftime('%Y/%m/%d')
        
        def _lot_sort_key(lot):
            procs = lot.get('processes', [])
            
            is_on_hold = any(
                p['status'] == 'current' and p.get('hold_reason') and p['hold_reason'] not in [None, '', 'None']
                for p in procs
            )
            is_single = len(procs) == 1 and not is_on_hold
            
            future_dates = [p['date'] for p in procs if p['status'] == 'future']
            all_dates = [p['date'] for p in procs]
            last_date = max(all_dates) if all_dates else ''
            last_future = max(future_dates) if future_dates else ''
            
            if is_single:
                # Group 4: abnormal - last
                return (4, last_date)
            elif is_on_hold:
                # Group 3: on-hold
                return (3, last_date)
            elif future_dates and last_future <= soon_cutoff:
                # Group 0: completing within 2 weeks - sort by completion asc
                return (0, last_future)
            elif future_dates:
                # Group 1: completing later - sort by completion asc
                return (1, last_future)
            else:
                # Group 2: all done - sort by last date (will reverse later)
                return (2, last_date)
        
        # Sort by group asc, then handle date direction per group
        self.lots_data.sort(key=_lot_sort_key)
        
        # Reverse date order for groups 2, 3, 4 (desc) while keeping 0, 1 (asc)
        # Use stable sort: split, reverse completed/hold/abnormal, recombine
        group_0_1 = [l for l in self.lots_data if _lot_sort_key(l)[0] in (0, 1)]
        group_2 = [l for l in self.lots_data if _lot_sort_key(l)[0] == 2]
        group_3 = [l for l in self.lots_data if _lot_sort_key(l)[0] == 3]
        group_4 = [l for l in self.lots_data if _lot_sort_key(l)[0] == 4]
        
        group_2.reverse()
        group_3.reverse()
        group_4.reverse()
        
        self.lots_data = group_0_1 + group_2 + group_3 + group_4
        self._group_counts = {
            'soon': len([l for l in group_0_1 if _lot_sort_key(l)[0] == 0]),
            'later': len([l for l in group_0_1 if _lot_sort_key(l)[0] == 1]),
            'completed': len(group_2),
            'on_hold': len(group_3),
            'abnormal': len(group_4),
        }
        self.page_size = page_size
        self.cell_width = cell_width
        self.show_all_labels = show_all_labels
        self.current_page = 0
        
        # Auto-detect date range from data
        all_dates = []
        for lot in self.lots_data:
            for proc in lot.get('processes', []):
                all_dates.append(proc['date'])
        
        self.start_date = start_date if start_date else (min(all_dates) if all_dates else None)
        self.end_date = end_date if end_date else (max(all_dates) if all_dates else None)
        
        self.dates = self._generate_dates()
        
        if container:
            with container:
                self._create_paginated_gantt()
        else:
            self._create_paginated_gantt()
    
    def _generate_dates(self, date_format: str = '%Y/%m/%d') -> List[str]:
        """Generate list of dates between start and end."""
        if not self.start_date or not self.end_date:
            return []
        start = datetime.strptime(self.start_date, date_format)
        end = datetime.strptime(self.end_date, date_format)
        dates = []
        current = start
        while current <= end:
            dates.append(current.strftime(date_format))
            current += timedelta(days=1)
        return dates
    
    def _create_paginated_gantt(self):
        """Create the Gantt chart with pagination controls."""
        total_pages = max(1, (len(self.lots_data) + self.page_size - 1) // self.page_size)
        
        with ui.row().classes('w-full items-center justify-between mb-2'):
            gc = self._group_counts
            self.info_label = ui.label(
                f'合計仕掛数: {len(self.lots_data)} 台 | '
                f'14日以内に完成予定: {gc["soon"]} 台 | '
                f'14日以降に完成予定: {gc["later"]} 台 | '
                f'保留: {gc["on_hold"]} 台 | '
                f'スケジューラ異常: {gc["abnormal"]} 台'
            ).style('font-size: 12px;')
            with ui.row().classes('items-center gap-2'):
                self.prev_btn = ui.button('◀ Prev', on_click=self._prev_page).props('flat dense')
                self.page_label = ui.label(f'Page {self.current_page + 1} / {total_pages}')
                self.next_btn = ui.button('Next ▶', on_click=self._next_page).props('flat dense')
        
        # Scrollable container — max-width stops it from growing beyond viewport
        self.gantt_container = ui.element('div').classes(f'{self._uid} border rounded').style(
            'overflow-x: auto; overflow-y: auto; max-height: 500px; '
            'max-width: calc(100vw - 100px); width: 100%;'
        )
        
        with self.gantt_container:
            self._render_current_page()
        
        self._update_pagination_buttons()
    
    def _get_page_dates(self, page_lots: List[dict]) -> List[str]:
        """Get only dates that have actual process entries on this page."""
        page_dates = set()
        for lot in page_lots:
            for proc in lot.get('processes', []):
                page_dates.add(proc['date'])
        
        if not page_dates:
            return []
        
        # Return sorted unique dates that have data (no gap-filling)
        return sorted(page_dates)

    def _render_current_page(self):
        """Render the current page of lots."""
        self.gantt_container.clear()
        
        start_idx = self.current_page * self.page_size
        end_idx = min(start_idx + self.page_size, len(self.lots_data))
        page_lots = self.lots_data[start_idx:end_idx]
        
        # Compute per-page date range
        page_dates = self._get_page_dates(page_lots)
        
        html_content = self._generate_gantt_html_for_lots(page_lots, page_dates)
        
        with self.gantt_container:
            ui.html(html_content, sanitize=False)
        
        # Scroll to today if present, otherwise center on latest date
        today_str = datetime.now().strftime('%Y/%m/%d')
        has_today = today_str in page_dates
        self._scroll_to_target(has_today)
    
    def _scroll_to_target(self, has_today: bool):
        """Scroll to today's column if present, otherwise center on the latest date."""
        if has_today:
            selector = "[data-today='true']"
        else:
            selector = "[data-latest='true']"
        
        ui.run_javascript(f'''
            setTimeout(() => {{
                const container = document.querySelector(".{self._uid}");
                if (!container) return;
                const target = container.querySelector("{selector}");
                if (target) {{
                    const scrollTarget = target.offsetLeft - (container.clientWidth / 2) + ({self.cell_width} / 2);
                    container.scrollLeft = Math.max(0, scrollTarget);
                }}
            }}, 200);
        ''')
    
    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render_current_page()
            self._update_pagination_buttons()
    
    def _next_page(self):
        total_pages = (len(self.lots_data) + self.page_size - 1) // self.page_size
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._render_current_page()
            self._update_pagination_buttons()
    
    def _update_pagination_buttons(self):
        total_pages = max(1, (len(self.lots_data) + self.page_size - 1) // self.page_size)
        self.page_label.set_text(f'Page {self.current_page + 1} / {total_pages}')
        self.prev_btn.set_enabled(self.current_page > 0)
        self.next_btn.set_enabled(self.current_page < total_pages - 1)
    
    def _generate_gantt_html_for_lots(self, lots: List[dict], page_dates: List[str] = None) -> str:
        """Generate HTML for a subset of lots with only relevant date columns."""
        cell_width = self.cell_width
        today_str = datetime.now().strftime('%Y/%m/%d')
        
        # Use page-specific dates if provided, otherwise fall back to full range
        dates = page_dates if page_dates else self.dates
        latest_date = dates[-1] if dates else None
        
        lot_col_width = 120
        total_width = lot_col_width + (cell_width * len(dates))
        
        styles = f"""
        <style>
            .gantt-table {{ 
                border-collapse: collapse; 
                font-family: Arial, sans-serif; 
                font-size: 12px;
                white-space: nowrap;
                table-layout: fixed;
                width: {total_width}px;
                min-width: {total_width}px;
            }}
            .gantt-table th, .gantt-table td {{
                border: 1px solid #ddd; padding: 0; text-align: center;
                height: 45px; width: {cell_width}px; min-width: {cell_width}px;
                max-width: {cell_width}px;
            }}
            .gantt-table th {{ 
                background: #f5f5f5; font-weight: normal; padding: 4px 2px; font-size: 10px;
                position: sticky; top: 0; z-index: 2;
            }}
            .lot-cell {{ 
                width: {lot_col_width}px !important; min-width: {lot_col_width}px !important; 
                max-width: {lot_col_width}px !important;
                padding: 8px; font-weight: bold; background: #fafafa; 
                text-align: left; font-size: 11px;
                position: sticky; left: 0; z-index: 1;
                overflow: hidden; text-overflow: ellipsis;
            }}
            .gantt-table thead th.lot-cell {{
                z-index: 3;
            }}
            .today-col {{
                background: #e3f2fd !important; color: #1565C0 !important; font-weight: bold !important;
            }}
            .today-cell {{
                background: #f5f9ff;
            }}
            .progress-cell {{ position: relative; padding: 0; }}
            .done {{
                position: absolute; top: 50%; left: 0; right: 0;
                height: 3px; background: #4CAF50; transform: translateY(-50%);
            }}
            .done-label {{
                position: absolute; bottom: 2px; left: 50%; transform: translateX(-50%);
                color: #4CAF50; font-size: 8px; white-space: nowrap; max-width: {cell_width - 4}px;
                overflow: hidden; text-overflow: ellipsis;
            }}
            .current {{
                position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
                width: 10px; height: 10px; background: #2196F3; border-radius: 50%;
                border: 2px solid #fff; box-shadow: 0 0 4px rgba(0,0,0,0.3);
            }}
            .current-label {{
                position: absolute; top: 2px; left: 50%; transform: translateX(-50%);
                color: #2196F3; font-size: 8px; font-weight: bold; white-space: nowrap;
                max-width: {cell_width - 4}px; overflow: hidden; text-overflow: ellipsis;
            }}
            .current-line {{
                position: absolute; top: 50%; left: 0; right: 50%;
                height: 3px; background: #2196F3; transform: translateY(-50%);
            }}
            .current-future {{
                position: absolute; top: 50%; left: 50%; right: 0;
                height: 0; border-top: 2px dashed #bbb; transform: translateY(-50%);
            }}
            .future {{
                position: absolute; top: 50%; left: 0; right: 0;
                height: 0; border-top: 2px dashed #bbb; transform: translateY(-50%);
            }}
            .future-label {{
                position: absolute; top: 2px; left: 50%; transform: translateX(-50%);
                color: #888; font-size: 8px; white-space: nowrap;
                max-width: {cell_width - 4}px; overflow: hidden; text-overflow: ellipsis;
            }}
            .hold-marker {{
                position: absolute; bottom: 4px; right: 2px;
                color: #f44336; font-size: 10px;
            }}
            .single {{
                position: absolute; top: 50%; left: 0; right: 0;
                height: 6px; background: #f44336; transform: translateY(-50%);
                border-radius: 2px;
            }}
            .single-label {{
                position: absolute; bottom: 2px; left: 50%; transform: translateX(-50%);
                color: #f44336; font-size: 8px; font-weight: bold; white-space: nowrap;
                max-width: {cell_width - 4}px; overflow: hidden; text-overflow: ellipsis;
            }}
        </style>
        """
        
        # Header row with today highlight + data attributes for JS scroll
        header_cells = ''
        for d in dates:
            attrs = ''
            classes = []
            if d == today_str:
                classes.append('today-col')
                attrs += " data-today='true'"
            if d == latest_date:
                attrs += " data-latest='true'"
            cls = f' class="{" ".join(classes)}"' if classes else ''
            header_cells += f'<th{cls}{attrs}>{d}</th>'
        header = f'<thead><tr><th class="lot-cell">Lot</th>{header_cells}</tr></thead>'
        
        # Body rows
        rows = []
        for lot in lots:
            lot_number = lot.get('lot_number', '')
            processes = lot.get('processes', [])
            
            # Detect single-process lot (not on hold)
            is_single = len(processes) == 1 and not (
                processes[0].get('hold_reason') and processes[0]['hold_reason'] not in [None, '', 'None']
            )
            
            cells = []
            
            # Build a map of which dates have data for this lot
            proc_by_date = {}
            for proc in processes:
                date = proc['date']
                if date not in proc_by_date:
                    proc_by_date[date] = []
                proc_by_date[date].append(proc)
            
            # Find the date indices that have actual processes for this lot
            lot_process_dates = sorted(set(proc['date'] for proc in processes))
            first_date = lot_process_dates[0] if lot_process_dates else None
            last_date = lot_process_dates[-1] if lot_process_dates else None
            
            # For gap filling: determine status before and after each gap
            # Build ordered list of (date, status) for this lot - prioritize current > done > future
            date_status_map = {}
            status_priority = {'current': 2, 'done': 1, 'future': 0}
            for proc in processes:
                d = proc['date']
                existing = date_status_map.get(d)
                if existing is None or status_priority.get(proc['status'], 0) > status_priority.get(existing, 0):
                    date_status_map[d] = proc['status']
            
            for date in dates:
                procs = proc_by_date.get(date, [])
                is_today = date == today_str
                today_class = ' today-cell' if is_today else ''
                
                # Check if this date is in a gap between processes
                in_range = first_date and last_date and first_date <= date <= last_date
                
                if not procs and not in_range:
                    # Outside lot's range — empty
                    cells.append(f'<td class="progress-cell{today_class}"></td>')
                    continue
                
                if not procs and in_range:
                    # Gap between processes — find the preceding status
                    prev_status = None
                    for pd in sorted(date_status_map.keys()):
                        if pd < date:
                            prev_status = date_status_map[pd]
                        else:
                            break
                    
                    if prev_status == 'done':
                        cell_content = '<div class="done"></div>'
                    elif prev_status == 'current':
                        cell_content = '<div class="future"></div>'
                    else:
                        cell_content = '<div class="future"></div>'
                    
                    cells.append(f'<td class="progress-cell{today_class}">{cell_content}</td>')
                    continue
                
                main_proc = procs[-1]
                
                # Determine cell status - prioritize: current > done > future
                if any(p['status'] == 'current' for p in procs):
                    status = 'current'
                elif any(p['status'] == 'done' for p in procs):
                    status = 'done'
                else:
                    status = main_proc.get('status', 'empty')
                
                # Pick label from the highest-priority process
                current_procs = [p for p in procs if p['status'] == 'current']
                labels = [p['label'] for p in procs if p.get('label')]
                if current_procs:
                    label = current_procs[-1]['label']
                else:
                    label = labels[-1] if labels else ''
                tooltip = ' | '.join(labels) if len(labels) > 1 else label
                
                has_hold = any(
                    p.get('hold_reason') and p['hold_reason'] not in [None, '', 'None'] 
                    for p in procs
                )
                
                cell_content = ''
                if is_single:
                    cell_content = '<div class="single"></div>'
                    if label:
                        cell_content += f'<div class="single-label" title="{tooltip}">{label}</div>'
                elif status == 'done':
                    cell_content = '<div class="done"></div>'
                    if label and self.show_all_labels:
                        cell_content += f'<div class="done-label" title="{tooltip}">{label}</div>'
                    if has_hold:
                        cell_content += '<div class="hold-marker">⚠</div>'
                elif status == 'current':
                    cell_content = '<div class="current-line"></div>'
                    cell_content += '<div class="current"></div>'
                    cell_content += '<div class="current-future"></div>'
                    if label:
                        cell_content += f'<div class="current-label" title="{tooltip}">{label}</div>'
                    if has_hold:
                        cell_content += '<div class="hold-marker">⚠</div>'
                elif status == 'future':
                    cell_content = '<div class="future"></div>'
                    if label:
                        cell_content += f'<div class="future-label" title="{tooltip}">{label}</div>'
                
                cells.append(f'<td class="progress-cell{today_class}">{cell_content}</td>')
            
            rows.append(f'<tr><td class="lot-cell" title="{lot_number}">{lot_number}</td>{"".join(cells)}</tr>')
        
        body = f'<tbody>{"".join(rows)}</tbody>'
        
        return f'{styles}<table class="gantt-table">{header}{body}</table>'

    def parse_lot_data(self, lot_data: dict) -> list[dict]:
        """
        Parse lot data dictionary into the format needed for Gantt chart.
        
        Args:
            lot_data: Dict{ lot_id: List[ Tuple(ProcessName, Timestamp, HoldReason, LeadTime) ] }
                    - LeadTime = 0: executed/done
                    - LeadTime > 0 or nan: scheduled/future
        
        Returns:
            List of lot dictionaries with processes
        """
        lots = []
        
        for lot_id, processes in lot_data.items():
            sorted_processes = sorted(processes, key=lambda x: x[1])
            
            last_done_idx = -1
            for i, proc in enumerate(sorted_processes):
                lead_time = proc[3]
                if lead_time == 0:
                    last_done_idx = i
            
            # Find the current process date
            current_date = None
            #current_timestamp = None
            for i, proc in enumerate(sorted_processes):
                if proc[3] == 0 and i == last_done_idx:
                    ts = proc[1]
                    if hasattr(ts, 'strftime'):
                        current_date = ts.strftime('%Y/%m/%d')
                    else:
                        current_date = str(ts)[:10].replace('-', '/')
                    #current_timestamp = ts
            
            # Build process list with status, filtering out futures before current
            process_list = []
            for i, proc in enumerate(sorted_processes):
                process_name, timestamp, hold_reason, lead_time = proc
                
                if lead_time == 0:
                    if i == last_done_idx:
                        status = 'current'
                    else:
                        status = 'done'
                else:

                    status = 'future'
                    #if current_timestamp:
                    #    if isnan(lead_time):
                    #        lead_time = 0
                    #    current_timestamp += timedelta(minutes=int(lead_time))
                    #    timestamp = current_timestamp
                if hasattr(timestamp, 'strftime'):
                    date_str = timestamp.strftime('%Y/%m/%d')
                else:
                    date_str = str(timestamp)[:10].replace('-', '/')
                
                # Skip future processes that are before or on the current process date
                if status == 'future' and current_date and date_str <= current_date:
                    continue
                
                process_list.append({
                    'date': date_str,
                    'status': status,
                    'label': process_name,
                    'hold_reason': hold_reason,
                    'lead_time': lead_time
                })
            
            lots.append({
                'lot_number': str(lot_id),
                'processes': process_list
            })
        
        return lots