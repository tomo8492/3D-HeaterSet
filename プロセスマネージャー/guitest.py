import os, sys, logging, queue
from nicegui import ui, events, app, Client
from datetime import datetime, date
import xception as E, module as M, pandas as pd, database_op as _DO, random, json
from typing import Dict, List, Tuple, Literal, Callable
from aggrid_jp import japanese_locale_text
from functools import partial, wraps
from collections import Counter
from sys_control import set_up_env, get_frozen_stat
from customwidget import CustomDialog, CustomEditor, CustomMenu, CustomTable, CustomTabPanel, CustomExpansion
import configparser, traceback, copy, re, asyncio
import schedule_utils as SU
from modules.xw_excel import append_reschedule_row, wait_for_book_close
from gantt_dialog import _GANTT_CSS, open_gantt_dialog


set_up_env()

_app_logger = logging.getLogger('guitest')

#MARK: Configuration
APP_NAME = "プロセスマネージャー"
TP_APP_NAME = "伊勢原工場倉庫係システム"
VERSION_NUMBER_MAIN = "0.62.4"
VERSION_NUMBER_TP = "0.1.2"
APP_VERSION = f"{VERSION_NUMBER_MAIN} beta "
TP_APP_VERSION = f"{VERSION_NUMBER_TP} beta "
TEST = frozen = get_frozen_stat()
DATAVIEWER_MODE_STRING = 'SKDF'
ACCENT_COLOR = '#569bbe'
SUB_ACCENT_COLOR = "#ffffff"

TAB_INITIALSED = True






WARNING_NOTIFICATION = {
    'NoKey' : ('warning', '有効な値はありません'),
    'IntErr' : ('warning', 'エラーが発生しました'),
    'NoData' : ('negative', '検索した値に関連するデータはありません。'),
    'UnExp' : ('warning', 'エラーが発生しました：')
}

DEFAULT_COLUMNS_SETTINGS = {
            'sortable' : False,
            'autoSize' : False,
            'flex' : 1
            
        }

UNIVERSAL_GRID_SETTING = {
                            'rowSelection': 'multiple'
                            }

#AGGRID_SELECTION: dict[str, Dict[str, 'ColumnVisbility']] = {}

CLICK_EVENT_CACHE: events.GenericEventArguments | None = None

LIKE_SEARCH = {}

MEMO_STORAGE = {}

MEMO_OBJ = None

TP_SHELF_MST = {}

QUERY_QUEUE = queue.Queue()

RESCHEDULE_EXCEL_CONFIG = {
    #
    'path':  r'ms-excel:ofe|u|https://nhkspggp.sharepoint.com/sites/PJ_265/%E4%BC%8A%E5%8B%A2%E5%8E%9F%E5%B7%A5%E5%A0%B4%E6%A5%AD%E5%8B%99%E8%AA%B2%E3%82%B5%E3%82%A4%E3%83%88/Shared%20Documents/%E3%83%AA%E3%82%B9%E3%82%B1%E4%BE%9D%E9%A0%BC/%E3%83%AA%E3%82%B9%E3%82%B1%E4%BE%9D%E9%A0%BC.xlsm?web=1',  # TODO: set before deployment
    'sheet': 'リスケ依頼表',  # TODO: confirm sheet name before deployment
}



def version_check() -> bool:
    config = configparser.ConfigParser()
    try:
        config.read(r"\\sflise02\伊勢原２\接合・セラミック部\業務課\2.業務G\Apps\プロセスマネージャー\config.ini", encoding='utf-8')
    except:
        raise Exception('ini設定ファイルが見つかりません、インターネットの接続と、ファイルの所在を確認してください。')
        
    latest = config['version']['ver']
    latest = int(latest.replace('.', ''))
    if latest > int(VERSION_NUMBER_MAIN.replace('.', '')):
        return False
    else:
        return True

#MARK: KeyBoard Event Cache
class keyCache:
    def __init__(self):
        self.key = None
        self.ctrl = False
        self.alt = False
        self.shift = False

    def update(self, event: events.KeyEventArguments):
        self.key = str(event.key).lower()
        self.ctrl = event.modifiers.ctrl
        self.alt = event.modifiers.alt
        self.shift = event.modifiers.shift

    def reset(self):
        self.key = None
        self.ctrl = False
        self.alt = False
        self.shift = False

    def __repr__(self):
        return f'key = {self.key} modifiers = ctrl({self.ctrl}) alt = ({self.alt}) shift = ({self.shift})'
    
def key_listener(event: events.KeyEventArguments):
    global KEY_CACHE
    KEY_CACHE.update(event)

def key_listener_tester():
    print(KEY_CACHE)

def click_cache(event):
    global CLICK_EVENT_CACHE
    CLICK_EVENT_CACHE = None
    CLICK_EVENT_CACHE = event

def get_click_cache() -> events.GenericEventArguments | None:
    if not isinstance(CLICK_EVENT_CACHE, events.GenericEventArguments):
        notification('negative', '選択中の行はありません。')
    return CLICK_EVENT_CACHE


#MARK: User Instance
class UserInstance():
    def __init__(self):
        self.uid = os.getlogin()
        self.user_name: str | None = None
        self.log = []

    def retry_log(self, target: Callable, *args):
        retry_obj = (target, args)
        self.log.append(retry_obj)

    def initial_user(self):
        user_settings = _DO.init_user_data(self.uid)
        self.user_name = user_settings.get('user_name') or None
        memo: Tuple[_DO.Memo | None] = user_settings.get('memo', ())
        if memo:
            for item in memo:
                MEMO_STORAGE[item.memoid] = (item.memocontent, item.memotitle)
        setting: _DO.Setting | None = user_settings.get('setting', None)
        if setting:
            setting_dict = json.loads(setting.settingcluster)
            if 'tag' in setting_dict:
                global TAG_COLOR
                TAG_COLOR = setting_dict['tag']
            if 'wip' in setting_dict:
                global wip_option
                wip_option = setting_dict['wip']

#MARK: Other Configuration Const
USER_INS = None

KEY_CACHE = keyCache()

ON_TYPE_SEARCH_ALERT = False

TAG_COLOR = {
    '赤' : '#FF0101',
    '緑' : '#10F015',
    '青' : '#047AFC',
    '黄' : '#FFF901',
    'オレンジ' : '#F6BE0A',
    '紫' : '#F802FE'
}


END_OF_BATCH: bool | None = None

SPINNER_CONTAINER: ui.dialog | None = None
SPINNER_TASK_REF: dict = {'task': None, 'cancel_btn': None, 'elapsed_label': None}

current_tabs: Dict[str, Tuple[Tuple[str, int], ui.tab, ui.tab_panel]] = {}



#MARK: Settings Const
search_mode = {'value' : 'i', 'lot': True}
wip_option = {'ise' : True, 'miya' : False, 'prestart' : False}
search_prompt = {'prompt' : True}

def spinner_create():
    global SPINNER_CONTAINER, SPINNER_TASK_REF
    with ui.dialog().props('persistent backdrop-filter="blur(5px) brightness(60%)"') as sp_container:
        with ui.card().classes('items-center gap-2').style('min-width:120px;'):
            ui.spinner(size='lg')
            cancel_btn = ui.button('キャンセル', icon='cancel').props('color=negative dense').classes('w-full').style('display:none;')
    SPINNER_CONTAINER = sp_container
    SPINNER_TASK_REF['cancel_btn'] = cancel_btn
    SPINNER_TASK_REF['elapsed_label'] = None



#MARK: Tags and Spinner functions
def color_tag_rules():
    holder = {}
    #for tag in TAG_COLOR.keys():
    #    holder[f':!bg-[{TAG_COLOR[tag]}]'] = f'(params) => params.data.Selected == "{tag}"'

    for tag, val in TAG_COLOR.items():
        holder[f':!bg-[{val}]'] = f'(params) => params.data.Selected == "{tag}"'
    return holder

def spinner_open():
    SPINNER_CONTAINER.open()


def spinner_close():
    SPINNER_CONTAINER.close()

def spinner_deco(func):
    import asyncio as _asyncio
    @wraps(func)
    async def wrapper(*args, **kwargs):
        spinner_open()
        cancel_btn = SPINNER_TASK_REF.get('cancel_btn')
        elapsed_label = SPINNER_TASK_REF.get('elapsed_label')
        tick_task = None

        # Store the *current* task so the cancel button can cancel it.
        # Do NOT use ensure_future for func — that would strip NiceGUI's slot
        # stack context and break any UI creation inside the decorated function.
        current = _asyncio.current_task()
        SPINNER_TASK_REF['task'] = current

        def _on_cancel():
            if current and not current.done():
                current.cancel()

        if cancel_btn:
            cancel_btn.on_click(_on_cancel)

        async def _tick():
            elapsed = 0
            while True:
                await _asyncio.sleep(1)
                elapsed += 1
                if elapsed >= 30 and cancel_btn:
                    cancel_btn.style('display:block;')

        tick_task = _asyncio.ensure_future(_tick())
        try:
            await func(*args, **kwargs)
        except _asyncio.CancelledError:
            _app_logger.warning(f'spinner_deco: task cancelled — {func.__name__}')
            notification('warning', f'処理がキャンセルされました: {func.__name__}')
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            _app_logger.error(tb)
            notification('warning', f'エラーが発生しました: {e}')
        finally:
            tick_task.cancel()
            SPINNER_TASK_REF['task'] = None
            if cancel_btn:
                cancel_btn.style('display:none;')
            spinner_close()
    return wrapper
    



#MARK: ColumnVisbility Class And Functions
class ColumnVisbility:
    def __init__(self, fileldid: str, colid: int):
        self.filedid = fileldid
        self.colid = colid
        self.visible = True
        self.highlight = False

    def flick(self, set_to: bool = None):
        if set_to is not None:
            self.visible = set_to
            return
        #print(f'before = {self.visible}')
        self.visible = not self.visible
        #print(f'after = {self.visible}')

    def highlight_flip(self, set_to: bool = None):
        if set_to is not None:
            self.highlight = set_to
            return
        self.highlight = not self.highlight
    
    def __repr__(self):
        return f'-~~((filedid = {self.filedid}\ncolid = {self.colid}\nvisibility = {self.visible}))~~-\n\n'

#def column_select(grid: 'AgGrid', mode: str | List[str] | Tuple[str] = 'all_on'):
#    columns = grid.selection
#    if mode == 'all_on':
#        for col in columns.values():
#            col.visible = True
#    if mode == 'all_off':
#        for col in columns.values():
#            col.visible = False
#    if isinstance(mode, (list, tuple)):
#        for col in columns.items():
#            col[1].visible = True



#MARK: Notification Function
def notification(type: Literal[
               'positive',
               'negative',
               'warning',
               'info',
               'ongoing',
           ], message: str, postion = 'bottom', close_button = False, color = None, multiline = False):
    ui.notify(message= message,
              type= type,
              position= postion,
              close_button= close_button,
              color= color,
              multi_line= multiline)
class AgGrid():
    def __init__(self):
        self.grid = None
        self.grid_container = None
        self.indexed_col_header = None
        self.column_list: Dict[str, ColumnVisbility] = {}
        self.options = {'allow_select_all': True,
                        }
        self.aggrid_col_setting = None
        #self.selection: Dict[str, ColumnVisbility] = {}
        
    def init_indexed_col_header(self, mode: str) -> dict:
        from aggridsetting import generate_col_settings
        self.aggrid_col_setting = aggrid_column_setting = generate_col_settings(mode)
        indexed_column_header = {headername['field'] : (index, headername['headerName']) for index, headername in enumerate(
            col for col in aggrid_column_setting 
        )}
        self.indexed_col_header = indexed_column_header
        return indexed_column_header
    
    def init_column_settings(self, mode:str, column_check: bool = True) -> Tuple[ui.card, List[ui.skeleton]] | None:
        ich = self.init_indexed_col_header(mode)
        if column_check:
            if mode in DATAVIEWER_MODE_STRING:
                for col, val in ich.items():
                    if col == 'Selected':
                        continue
                    column_data: Tuple[int, str] = val
                    new_col_object = ColumnVisbility(col, column_data[0])
                    self.column_list[col] = new_col_object
            else:
                with ui.row().classes('w-full gap-x-4') as col_check_container:
                    skeleton = ui.skeleton('rect').classes('w-full h-8')
                    sk_list = [skeleton]
                    skeletons_up(sk_list, True)
                    with ui.card().classes('flex-1') as col_check_card:
                        with ui.row():
                            for col, val in ich.items():
                                if col == 'Selected':
                                    continue
                                column_data: Tuple[int, str] = val
                                new_col_object = ColumnVisbility(col, column_data[0])

                                newcheck = ui.checkbox(column_data[1]).bind_value(new_col_object, 'visible')
                                #print(f'{column_data[1]} binded to {new_col_object}')
                                self.column_list[col] = new_col_object
                col_check_card.set_visibility(False)
                return col_check_card, sk_list

    async def awaitrowdata(self):
            grid = self.grid
            rows = await grid.get_selected_rows()
            selected = self.get_selected_cols(rows)
            if rows:
                if M.copytoclipboard(selected):
                    notification('positive', '行をコピーしました。')

    async def _selectallrows(self):
        grid = self.grid
        await grid.run_grid_method('selectAll','filtered')
        #global END_OF_BATCH
        #filtered_model = await grid.run_grid_method('getFilterModel')
        #if not filtered_model:
        #    await grid.run_grid_method("(grid) => grid.selectAll('filtered' )")
        #else:
        #    selected = await grid.get_client_data(method='filtered_unsorted')
        #    selected_id = [row['lotid'] for row in selected]
        #    END_OF_BATCH = False
        #    for id in selected_id:
        #        grid.run_row_method(id, 'setSelected', True)
        #    END_OF_BATCH = True

    def _deselectallrows(self):
            self.grid.run_grid_method('deselectAll', 'filtered')
    
            

    def _deselect(self):
        grid = self.grid
        for col in grid.options['columnDefs']:
            col['headerClass'] = 'None'
        self.deselect_all()
        aggrid_update(grid, grid.options) 

    def copy_cell(self):
        global CLICK_EVENT_CACHE
        if CLICK_EVENT_CACHE:
            if M.singlecopy(CLICK_EVENT_CACHE.args['value']):
                CLICK_EVENT_CACHE = None
                notification('positive', 'セルをコピーしました。')

    async def keyboard_copy(self):
        if KEY_CACHE.key == 'c' and KEY_CACHE.ctrl:
            await self.awaitrowdata()
        if KEY_CACHE.key == 'c' and KEY_CACHE.alt:
            self.copy_cell()
        if KEY_CACHE.key == 'a' and KEY_CACHE.ctrl:
            if self.options['allow_select_all']:
                await self._selectallrows()
            else:
                notification('info', 'このテーブルは大量なデータが入っており、アプリの動作異常を引き起こす恐れがあるため、全選択を禁止しています。')
                notification('info', 'フィルターを活用して必要なデータを絞ってコピーしてください。')


    def aggrid_creator(self, dataframe: pd.DataFrame, _classes: str = None, _table_view: bool = False, row_id: str = 'lotid') -> Tuple[ui.element, ui.aggrid]:
        with ui.element('div').classes('w-full h-full flex-grow overflow-hidden') as aggrid_container:
            # Create the AgGrid to display the data
            #print(aggrid_column_setting)
            columnsettings = self.aggrid_col_setting
            if columnsettings is None:
                raise E.InternalErrorExc('column setting is not defined.')
            grid = ui.aggrid.from_pandas(dataframe, auto_size_columns=False, options=
                        {
                'columnDefs' : copy.deepcopy(columnsettings),
                'defaultColDef' : {} if not _table_view else  {'width': 150, 'minWidth': 100},
                'rowSelection': {'mode': 'multiRow', 'selectAll': 'filtered', 'checkboxes': False, 'headerCheckbox': False, 'enableClickSelection': True},
                'enableSelectionWithoutKeys': False,
                'localeText': japanese_locale_text,
                'rowClassRules': color_tag_rules(),
                ':getRowId': f'(params) => params.data.{row_id}' if row_id else None,

            },
            theme='balham'
            ).classes('h-[calc(65vh-2rem)] overscroll-auto' if _classes is None else _classes)
        self.grid = grid
        self.grid_container = aggrid_container

        return aggrid_container, grid

    def highlight_column(self):
        if self.indexed_col_header is None:
            raise E.InternalErrorExc('indexed_col_header is not defined.')
        grid = self.grid
        global CLICK_EVENT_CACHE
        e = CLICK_EVENT_CACHE
        if not e:
            return
        col_name = e.args['colId']
        # Auto-populate column_list from indexed_col_header if the lazy menu
        # has not been opened yet (prevents KeyError on early Alt-click).
        if not self.column_list:
            for _col, _val in self.indexed_col_header.items():
                if _col == 'Selected':
                    continue
                self.column_list[_col] = ColumnVisbility(_col, _val[0])
        selection = self.column_list
        ind = self.indexed_col_header[col_name][0]
        altered_defs = dict(grid.options)
        attribute_target = altered_defs['columnDefs'][ind]
        #print(attribute_target)
        #print(grid.options)
        with grid._props.suspend_updates():
            if 'headerClass' not in attribute_target or attribute_target['headerClass'] != 'selected-column':
                attribute_target['headerClass'] = 'selected-column'
                selection[col_name].highlight_flip(True)
            elif attribute_target['headerClass'] == 'selected-column':
                attribute_target['headerClass'] = 'None'
                selection[col_name].highlight_flip(False)
        #grid.update()
            aggrid_update(grid, altered_defs)

    def left_click_event(self, event):
        click_cache(event)
        if KEY_CACHE.alt:
            self.highlight_column()

    def update_column(self):
        if self.column_list is None:
            raise E.InternalErrorExc('column_list is not defined.')
        grid = self.grid
        #for col in column_list.keys():
        #    grid.options['columnDefs'][column_list[col].colid]['hide'] = not column_list[col].visible
        for col in self.column_list.values():
            grid.options['columnDefs'][col.colid]['hide'] = not col.visible
        aggrid_update(grid, grid.options)
        #grid.update()

    def cellkey_func_bind(self):
        self.grid.on('cellKeyDown', self.keyboard_copy)
        self.grid.on('cellClicked', self.left_click_event)

    def get_selected_cols(self, rows: List[dict]):
        if not self.column_list:
            return tuple({k: v for k, v in row.items() if k != 'Selected'} for row in rows)
        if any(tuple(v.highlight for _, v in self.column_list.items() if self.column_list)):
            return tuple({k : v for k, v in row.items() if k != 'Selected' and self.column_list[k].visible and self.column_list[k].highlight} for row in rows)
        else:
            return tuple({k : v for k, v in row.items() if k != 'Selected' and self.column_list[k].visible} for row in rows)
    
    def deselect_all(self):
        for k in self.column_list.values():
            k.highlight_flip(False)

    def column_select(self, mode: str | List[str] | Tuple[str] = 'all_on'):
        columns = self.column_list
        if mode == 'all_on':
            for col in columns.values():
                col.visible = True
        if mode == 'all_off':
            for col in columns.values():
                col.visible = False
        if isinstance(mode, (list, tuple)):
            for col in columns.items():
                col[1].visible = True

    async def filter_api(self, column_name: str, filter_setting: Dict[str, str | List[Dict[str, str]]]):
        '''
        # Parameters
        column_name: The name of the column to be applied to. str

        filter_setting: A dictionary of filter type and conditions. dict

        # filter_setting
        sample: single condition = {'filterType': 'text', 'type': 'notEqual', 'filter': '-'}

        multiple condiction = {'filterType': 'text', 'operator': 'OR', 'conditions': [{'filterType': 'text', 'type': 'contains', 'filter': '支給'}, {'filterType': 'text', 'type': 'contains', 'filter': '梱包'}]}
        '''
        self.grid.run_grid_method('setFilterModel', 'null')
        self.grid.run_grid_method('setColumnFilterModel', column_name, filter_setting)
        self.grid.run_grid_method('onFilterChanged')

    async def apply_filter_model(self, filter_model_dict: dict, valid_columns: set):
        """Apply a full AG Grid filter model, silently skipping unknown columns."""
        safe_model = {k: v for k, v in filter_model_dict.items() if k in valid_columns}
        self.grid.run_grid_method('setFilterModel', safe_model if safe_model else 'null')
        self.grid.run_grid_method('onFilterChanged')


def datepicker_input(name: str = 'Date', defalut_date: str = None) -> Tuple[ui.row, ui.label]:
    with ui.row().classes('items-center') as row_container:
        #date_title = ui.label(name).classes('bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500')
        picker_obj = ui.icon('edit_calendar').classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg')
        date_label = ui.label('').classes('bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg')
        with ui.menu().props('no-parent-event') as menu:
            with ui.date().bind_value(date_label, 'text') as picker:
                if defalut_date:
                    picker.value = defalut_date
                with ui.row().classes('justify-end'):
                    ui.button('閉じる', on_click=menu.close).props('flat')
        picker_obj.on('click', menu.open)
            
    return row_container, date_label
            
#MARK: Query Related Functions
async def query_on_search(textinput: ui.textarea, skeleton: ui.skeleton, duplist: List[str]):
    data_input = textinput.value

    clean_key = M.validate_text(data_input, 'a', '5501', duplist) 
    for key, val in clean_key.items():
        cleankey_obj = val
        if len(key) >= 5:
            #MARK: SQL Search function 1
            result = await M.wrapped_sql_search(cleankey_obj.key, cleankey_obj.mode, plant = cleankey_obj.plant)
            #print(f'qos = {result}')
            if result != ('NoCacheAvaliable',):
                clean_key[key] = result
        else:
            clean_key[key] = ('DBのワイルドカード検索はできません、', 'DBのほかに少なくとも2桁の数字を入力してください。')

    return clean_key

async def query_data(textinput: ui.textarea, dialog: ui.dialog) -> Tuple[Dict[str, pd.DataFrame], M.query_]:
    #import asyncio
    #import time

    #start_of_query_process = time.perf_counter()
    data_input = textinput.value

    plant_list = []
    plant_str = ''
    if wip_option['ise']:
        plant_list.append('5501')
    if wip_option['miya']:
        plant_list.append('5401')
    
    if len(plant_list) == 2:
        plant_str = f'\'{plant_list[0]}\',\'{plant_list[1]}\''
    elif len(plant_list) == 1:
        plant_str = f'\'{plant_list[0]}\''
    else:
        warning_val = WARNING_NOTIFICATION['IntErr']
        notification(warning_val[0], warning_val[1] + f'工場コード不正: {plant_list}、Length: {len(plant_list)} Type: {type(plant_list)} Location: query_data')
        raise E.InternalErrorExc('Plant string not legit.')
    
    mode_ = None
    if search_mode['value'] == 'w' and not wip_option['prestart']:
        mode_ = 'e'
    else:
        mode_ = search_mode['value']

    try:
        clean_key = M.validate_text(data_input, mode_, plant_str)
    
    except E.NoKeyExc:
        warning_val = WARNING_NOTIFICATION['NoKey']
        notification(warning_val[0], warning_val[1])
        raise
    
    #print(f'Input validation completed in {time.perf_counter() - start_of_query_process:.2f} seconds')

    spinner_open()

    #start_of_query_execution = time.perf_counter()
    records = await uni_query(clean_key)
    #print(f'Query execution completed in {time.perf_counter() - start_of_query_execution:.2f} seconds')

    return records, clean_key

async def tp_query() -> Tuple[Dict[str, pd.DataFrame], M.query_]:
    clean_key = M.query_('', 'A', '5501')
    spinner_open()
    records = await uni_query(clean_key)
    return records
    

async def context_adapted_query() -> Dict[str, pd.DataFrame]:
    try:
        context = CLICK_EVENT_CACHE.args['colId']
        key = CLICK_EVENT_CACHE.args['data']['lotid']
    except Exception as e:
        notification('negative', f'予想外のエラーが発生しました,リトライしてください。：{e}')
        return
    if context == 'currentprocess':
        mode = 'v'
        result, _ = await query_data_stdin(mode, key)

        await create_process_dialog(schedule_= result)

    if context == 'lastupdatedate':
        mode = 'h'
        result, _ = await query_data_stdin(mode, key)
        #print(result)
        await create_process_dialog(history_= result)

    if context == 'restriction':
        mode = 'y'
        result, _ = await query_data_stdin(mode, key)

        await create_process_dialog(restriction_= result)

    if context == 'lotid':
        mode = 's'
        result, _ = await query_data_stdin(mode, key)
        #print(result)
        await create_process_dialog(serial_= result)



async def query_data_stdin(mode: str, key: str) -> Tuple[Dict[str, pd.DataFrame], M.query_]:
    
    if key is None:
        raise E.NoDataExc()

    plant_list = []
    plant_str = ''
    if wip_option['ise']:
        plant_list.append('5501')
    if wip_option['miya']:
        plant_list.append('5401')
    
    if len(plant_list) == 2:
        plant_str = f'\'{plant_list[0]}\',\'{plant_list[1]}\''
    elif len(plant_list) == 1:
        plant_str = f'\'{plant_list[0]}\''
    else:
        warning_val = WARNING_NOTIFICATION['IntErr']
        notification(warning_val[0], warning_val[1] + f'工場コード不正: {plant_list}、Length: {len(plant_list)} Type: {type(plant_list)} Location: query_data_stdin')
        raise E.InternalErrorExc('Plant string not legit.')
    
    mode_ = mode

    cond = (0, 0)
    if mode_ == 'v':
        cond = ("V_ASP_生産計画一覧.着手予定日", 'ASC')
    if mode_ == 'h':
        cond = ("MLH.UP_DATE","DESC")
    spinner_open()

    try:
        clean_key = M.validate_text(key, mode_, plant_str)
        #MARK:SQL Query funtion 2
        records = await M.wrapped_sql_search(clean_key.key, clean_key.mode, plant = clean_key.plant, cond = cond)
        dframe = M.validate_return(records, clean_key.mode)
    
    except E.NoKeyExc:
        warning_val = WARNING_NOTIFICATION['NoKey']
        notification(warning_val[0], warning_val[1])
        spinner_close()
        raise

    except E.NoDataExc:
        if mode_ in 'yvh':
            return {}, clean_key
        warning_val = WARNING_NOTIFICATION['NoData']
        notification(warning_val[0], warning_val[1])
        spinner_close()
        raise

    except Exception as e:
        warning_val = WARNING_NOTIFICATION['UnExp']
        notification(warning_val[0], f'{warning_val[1]}: {e}')
        spinner_close()
        raise

    return dframe, clean_key

async def query_data_stdin_no_validation(mode: str, key: List[str]) -> Tuple[Dict[str, pd.DataFrame], M.query_]:
    
    if not key:
        raise E.NoDataExc()

    plant_list = []
    plant_str = ''
    if wip_option['ise']:
        plant_list.append('5501')
    if wip_option['miya']:
        plant_list.append('5401')
    if len(plant_list) == 2:
        plant_str = f'\'{plant_list[0]}\',\'{plant_list[1]}\''
    elif len(plant_list) == 1:
        plant_str = f'\'{plant_list[0]}\''
    else:
        warning_val = WARNING_NOTIFICATION['IntErr']
        notification(warning_val[0], warning_val[1] + f'工場コード不正: {plant_list}、Length: {len(plant_list)} Type: {type(plant_list)} Location: query_data_stdin_no_validation')
        raise E.InternalErrorExc('Plant string not legit.')
    
    mode_ = mode

    clean_key = M.query_(key, mode_, plant_str)

    spinner_open()

    records = await uni_query(clean_key)
    
    return records, clean_key

async def query_whole_table_wo_vali(mode: str, _key: str = '') -> pd.DataFrame:
    spinner_open()
    try:
        #MARK:SQL Query funtion 3
        records = await M.wrapped_sql_search(_key, mode, plant = '5501', outaspd=True)
        if mode in 'KDS':
            records['index_id'] = range(len(records))
    except E.NoDataExc:
        notification('negative', 'データを検索できませんでした。')
        records = pd.DataFrame()
    except Exception as e:
        notification('warning', f'予想外のエラーが発生しました：\n{e}')
        spinner_close()
        raise

    spinner_close()
    return records

async def uni_query(clean_key: M.query_) -> pd.DataFrame | List[Tuple[str]]:
    #print(records) #!!
    try:
        #MARK:SQL Query funtion 4
        records = await M.wrapped_sql_search(clean_key.key, clean_key.mode, plant = clean_key.plant)
        dframe = M.validate_return(records, clean_key.mode)
        #print(f'uni_query result = {dframe}')
    except E.NoDataExc:
        warning_val = WARNING_NOTIFICATION['NoData']
        notification(warning_val[0], warning_val[1])
        spinner_close()
        raise
    except Exception as e:
        warning_val = WARNING_NOTIFICATION['NoData']
        notification(warning_val[0], warning_val[1] + str(e))
        spinner_close()
        raise
    
    return dframe

async def uni_query_tp(data: str, existing_data: List[str], mode: str = 'lot_status') -> Dict[str, str]:
    #spinner_open()
    raw_input = data
    records = None
    validate_mode = {
        'lot_status' : 'i'
    }
    try:
        clean_key = M.validate_text(raw_input, validate_mode.get(mode, 'i'), '5501')
        for k in clean_key.key:
            if k in existing_data:
                notification('negative', '重複した指図が入力されました、再度確認してください。')
                return None
        #MARK:SQL Query funtion 5
        records = await M.fast_query(mode, clean_key.key)
        #print(f'427 guitest -- > \n{records}')
        
    except E.NoKeyExc:
        notification('negative', '有効な指図はありません。')
        return None
    except E.NoDataExc:
        notification('negative', '関連データはありません、指図正しいかどうかを確認してください。', 'center')
        return None
    except Exception as e:
        notification('warning', f'予期しないエラーが発生しました：\n{e}')
    #print(records) #!!
    if not records:
        #spinner_close()
        notification('negative', '関連データはありません、指図正しいかどうかを確認してください。', 'center')
        raise E.NoDataExc()
    record_row = records[0]

    newrow = {'lotid': record_row[0], 'process': record_row[1], 'procode': record_row[3], 'status': record_row[2], 'shelf': record_row[4]}
    return newrow

#MARK: Helper Functions
async def quick_dialog_yes_no(message: str) -> int:
    with ui.dialog() as check_shelf, ui.card():
        ui.label(message).classes('text-lg')
        with ui.row():
            ui.button('はい', on_click=lambda: check_shelf.submit(1))
            ui.button('いいえ', on_click=lambda: check_shelf.submit(0))
        result = await check_shelf
    return result


#MARK: Excel Write Dialog
@spinner_deco
async def excel_write_dialog(data: list[dict]) -> None:
    from modules import ActiveBook, get_all_books_act

    if not data:
        notification('warning', 'データがありません。')
        return

    fields = [k for k in data[0].keys()]
    ab = ActiveBook()

    book_names = []
    active_book = None
    try:
        result = await get_all_books_act()
        if result:
            book_names, active_book = list(result[0]), result[1]
    except Exception:
        pass

    # Inline style constants for direction cards (active / inactive state)
    _CARD_ON  = 'flex:1; cursor:pointer; border:2px solid #37a660; border-radius:8px; background:#f0f9f4'
    _CARD_OFF = 'flex:1; cursor:pointer; border:2px solid #d1d5db; border-radius:8px; background:#ffffff'
    _LBL_ON   = 'font-size:14px; font-weight:600; color:#37a660'
    _LBL_OFF  = 'font-size:14px; font-weight:600; color:#6b7280'

    def _sec(title: str):
        """Green-accented section title."""
        with ui.row(align_items='center').classes('gap-2 mb-2'):
            ui.element('div').style(
                'width:3px; height:14px; background:#37a660; border-radius:3px; flex-shrink:0'
            )
            ui.label(title).style(
                'font-size:11px; font-weight:700; color:#37a660; '
                'text-transform:uppercase; letter-spacing:0.07em'
            )

    with ui.context.client.content:
        with ui.dialog().classes('w-full').props('persistent backdrop-filter="blur(8px) brightness(40%)"') as dialog:
            close_btn = ui.label('閉じる').classes('text-xl text-white/75 hover:text-white cursor-pointer')
            close_btn.on('click', dialog.close)
            with ui.card().classes('w-full p-0 gap-0 overflow-hidden').style('max-width:none'):

                # ── Header ───────────────────────────────────────────────
                with ui.row(align_items='center').classes('w-full justify-between px-5 py-3').style('background:#37a660'):
                    with ui.row(align_items='center').classes('gap-3'):
                        ui.icon('table_view').style('color:#fff; font-size:22px')
                        ui.label('Excel への書き出し').style('color:#fff; font-size:16px; font-weight:600')
                        ui.label(f'{len(data)} 行').style(
                            'color:#fff; font-size:11px; font-weight:500; '
                            'background:rgba(255,255,255,0.25); padding:2px 12px; border-radius:20px'
                        )
                    #ui.button(icon='close', on_click=dialog.close).props('flat round dense').style('color:#fff')

                # ── Scrollable body ───────────────────────────────────────
                with ui.scroll_area().classes('w-full').style('height:70vh'):

                    # 1. データプレビュー ──────────────────────────────────
                    with ui.column().classes('w-full gap-2 px-5 pt-4 pb-4').style('border-bottom:1px solid #f0f0f0'):
                        _sec('データプレビュー')
                        with ui.scroll_area().classes('w-full rounded-lg').style(
                            'height:108px; border:1px solid #e2e8f0; background:#f8fafb'
                        ):
                            preview_cols = [{'name': f, 'label': f, 'field': f, 'align': 'left'} for f in fields]
                            ui.table(columns=preview_cols, rows=data[:5]).classes('text-xs').props('dense flat')
                        ui.label(f'先頭 5 行を表示 — 全 {len(data)} 行が書き出されます').style(
                            'font-size:11px; color:#9ca3af'
                        )

                    # 2. 書き出し先 ───────────────────────────────────────
                    with ui.column().classes('w-full gap-3 px-5 pt-4 pb-4').style('border-bottom:1px solid #f0f0f0'):
                        _sec('書き出し先')
                        with ui.row().classes('w-full gap-3 flex-wrap item-center'):
                            with ui.column().classes('flex-1 gap-1').style('min-width:200px'):
                                ui.label('Excel ブック').style('font-size:11px; font-weight:600; color:#6b7280')
                                with ui.row(align_items='center').classes('w-full gap-1'):
                                    book_select = (
                                        ui.select(book_names, value=active_book)
                                        .classes('flex-1').props('outlined dense options-dense')
                                    )
                                    async def refresh_books():
                                        try:
                                            r = await get_all_books_act()
                                            if r:
                                                book_select.set_options(list(r[0]))
                                        except Exception as exc:
                                            notification('negative', f'ブック取得エラー：{exc}')
                                    book_select.on('popup-show', refresh_books)
                                    #ui.button(icon='refresh', on_click=refresh_books).props('flat round dense').classes('text-gray-400')

                            with ui.column().classes('flex-1 gap-1').style('min-width:150px'):
                                ui.label('シート').style('font-size:11px; font-weight:600; color:#6b7280')
                                sheet_select = ui.select([], label='').classes('w-full').props('outlined dense options-dense')

                        async def on_book_change(e):
                            if not e.value:
                                return
                            try:
                                await ab.set_target_book(e.value)
                                sheets = await ab.get_sheets()
                                sheet_select.set_options(sheets)
                            except Exception as ex:
                                notification('negative', f'シート取得エラー：{ex}')

                        book_select.on_value_change(on_book_change)

                        if active_book:
                            try:
                                await ab.set_target_book(active_book)
                                sheets = await ab.get_sheets()
                                sheet_select.set_options(sheets)
                            except Exception:
                                pass

                    # 3. 書き込み設定 ─────────────────────────────────────
                    with ui.column().classes('w-full gap-3 px-5 pt-4 pb-4').style('border-bottom:1px solid #f0f0f0'):
                        _sec('書き込み設定')

                        with ui.row().classes('gap-4 flex-wrap item-center'):
                            with ui.column().classes('gap-1'):
                                ui.label('開始列').style('font-size:11px; font-weight:600; color:#6b7280')
                                start_col = ui.input(
                                    placeholder='例: A',
                                    validation={'A–Z で入力してください': lambda v: bool(re.match(r'^[a-zA-Z]{1,3}$', v))},
                                ).classes('w-20').props('outlined dense')
                            with ui.column().classes('gap-1'):
                                ui.label('開始行').style('font-size:11px; font-weight:600; color:#6b7280')
                                start_row = ui.number(value=1, min=1).classes('w-24').props('outlined dense')

                        ui.label('書き出し方向').style('font-size:11px; font-weight:600; color:#6b7280')
                        direction = {'val': 'vertical'}
                        with ui.row().classes('w-full gap-2'):
                            with ui.card().props('flat').style(_CARD_ON) as vert_card:
                                with ui.row(align_items='center').classes('gap-3 p-1'):
                                    ui.icon('arrow_downward').style('color:#37a660; font-size:24px')
                                    with ui.column().classes('gap-0'):
                                        vert_lbl = ui.label('縦方向').style(_LBL_ON)
                                        ui.label('行ごとに下へ追加').style('font-size:11px; color:#9ca3af')
                            with ui.card().props('flat').style(_CARD_OFF) as horiz_card:
                                with ui.row(align_items='center').classes('gap-3 p-1'):
                                    ui.icon('arrow_forward').style('color:#9ca3af; font-size:24px')
                                    with ui.column().classes('gap-0'):
                                        horiz_lbl = ui.label('横方向').style(_LBL_OFF)
                                        ui.label('列ごとに右へ追加').style('font-size:11px; color:#9ca3af')

                        def set_vertical():
                            direction['val'] = 'vertical'
                            vert_card.style(_CARD_ON);   vert_lbl.style(_LBL_ON)
                            horiz_card.style(_CARD_OFF); horiz_lbl.style(_LBL_OFF)
                            lookup_ref.props('label="参照列（英字）"')
                            lookup_ref.validation = {
                                '参照列を入力してください': lambda v: bool(re.match(r'^[a-zA-Z]{1,3}$', v))
                            }

                        def set_horizontal():
                            direction['val'] = 'horizontal'
                            horiz_card.style(_CARD_ON);  horiz_lbl.style(_LBL_ON)
                            vert_card.style(_CARD_OFF);  vert_lbl.style(_LBL_OFF)
                            lookup_ref.props('label="参照行（数字）"')
                            lookup_ref.validation = {
                                '参照行を入力してください': lambda v: v.isdigit() and int(v) > 0
                            }

                        vert_card.on('click', set_vertical)
                        horiz_card.on('click', set_horizontal)

                        header_check = ui.checkbox('列ヘッダーを書き出す（データの 1 行上に列名を追加）').classes('mt-1')

                    # 4. ルックアップ設定 ──────────────────────────────────
                    with ui.column().classes('w-full gap-3 px-5 pt-4 pb-4'):
                        with ui.row(align_items='center').classes('w-full justify-between'):
                            with ui.row(align_items='center').classes('gap-3'):
                                ui.icon('manage_search').style('color:#6b7280; font-size:20px')
                                with ui.column().classes('gap-0'):
                                    ui.label('ルックアップモード').style('font-size:13px; font-weight:600; color:#374151')
                                    ui.label('Excel の既存データと照合して書き込む').style('font-size:11px; color:#9ca3af')
                            lookup_switch = ui.switch('').props('color=positive')

                        with ui.card().props('flat').style(
                            'width:100%; border:1.5px solid #d1e8dc; '
                            'background:#f8fafb; border-radius:8px'
                        ).bind_visibility_from(lookup_switch, 'value'):
                            with ui.column().classes('w-full gap-3 p-4'):
                                with ui.element('div').style(
                                    'background:#e8f5ee; border-left:3px solid #37a660; '
                                    'padding:8px 12px; border-radius:0 6px 6px 0'
                                ):
                                    ui.label(
                                        'Excel の参照列（またはヘッダー行）を検索し、'
                                        'データのキー列と一致するセルを見つけて書き込みます。'
                                    ).style('font-size:11px; color:#4b5563')
                                with ui.row().classes('gap-3 flex-wrap item-center'):
                                    with ui.column().classes('gap-1'):
                                        ui.label('参照列（Excel の列名）').style('font-size:11px; font-weight:600; color:#6b7280')
                                        lookup_ref = ui.input(
                                            placeholder='例: A',
                                            validation={'参照列を入力してください': lambda v: bool(re.match(r'^[a-zA-Z]{1,3}$', v))},
                                        ).classes('w-20').props('outlined dense')
                                    with ui.column().classes('flex-1 gap-1').style('min-width:150px'):
                                        ui.label('キー列（データ側の列）').style('font-size:11px; font-weight:600; color:#6b7280')
                                        key_field_select = ui.select(fields).classes('w-full').props('outlined dense options-dense')

                # ── Footer ───────────────────────────────────────────────
                with ui.row(align_items='center').classes('w-full justify-end gap-3 px-5 py-3').style(
                    'background:#f8fafb; border-top:1px solid #e5e7eb'
                ):
                    #ui.button('キャンセル', on_click=dialog.close).props('flat').classes('text-gray-500')

                    @spinner_deco
                    async def on_execute():
                        if not book_select.value:
                            notification('negative', 'ブックを選択してください。')
                            return
                        if not sheet_select.value:
                            notification('negative', 'シートを選択してください。')
                            return
                        if not re.match(r'^[a-zA-Z]{1,3}$', start_col.value or ''):
                            notification('negative', '開始列を正しく入力してください。')
                            return
                        if lookup_switch.value:
                            if direction['val'] == 'vertical':
                                if not re.match(r'^[a-zA-Z]{1,3}$', lookup_ref.value or ''):
                                    notification('negative', '参照列を正しく入力してください。')
                                    return
                            else:
                                lv = lookup_ref.value or ''
                                if not lv.isdigit() or int(lv) < 1:
                                    notification('negative', '参照行を正しく入力してください。')
                                    return
                            if not key_field_select.value:
                                notification('negative', 'キー列を選択してください。')
                                return
                        try:
                            row_val = int(start_row.value) if start_row.value else 1
                            if not lookup_switch.value:
                                count = await ab.simple_write(
                                    sheet_select.value, data, fields,
                                    start_col.value, row_val,
                                    direction['val'], header_check.value,
                                )
                            else:
                                count = await ab.lookup_write(
                                    sheet_select.value, data, fields,
                                    key_field_select.value, lookup_ref.value,
                                    start_col.value, row_val,
                                    direction['val'], header_check.value,
                                )
                            notification('positive', f'{count}行を書き出しました。')
                            dialog.close()
                        except (E.NoBookSet, E.NoBookFound):
                            notification('negative', 'ブックが選択されていません。')
                        except Exception as e:
                            notification('negative', f'書き出しエラー：{e}')

                    ui.button('書き出し実行', icon='play_arrow', on_click=on_execute).props('color=positive')

    dialog.open()


async def context_wrapper(func, *args, **kwargs):
    if CLICK_EVENT_CACHE is None:
        notification('negative', '選択中の行はありません。')
        return
    spinner_open()
    result = await func(*args, **kwargs)
    spinner_close()
    return result


def skeletons_before(skeleton_list: List[ui.skeleton]):
    def decorator(func: Callable):
        async def skeleton_wrapper(*args, **kwargs):
            skeletons_up(skeleton_list, True)
            result = func(*args, **kwargs)
            skeletons_up(skeleton_list, False)
            return result
        return skeleton_wrapper
    return decorator

def screen_element_copy(value: str):
    if not M.singlecopy(value):
        notification('negative', 'コピーできる値はありません。')
    else:
        notification('positive', 'コピーしました。')


        

#MARK: Drawer Related
def set_og_drawer(title:ui.label, _list:ui.list, vis: bool):
    title.set_visibility(vis)
    _list.set_visibility(vis)

def skeletons_up(ske_list: List[ui.skeleton], signal: bool):
    for ske in ske_list:
        ske.set_visibility(signal)

#MARK: Memo Drawer
def memo_drawer(title: ui.label, _list: ui.list, drawer: ui.left_drawer):
    global MEMO_OBJ
    global MEMO_STORAGE
    
    def close_memo():
        memo_list.set_visibility(False)
        set_og_drawer(title, _list, True)

    def active_placeholder():
        if len(MEMO_STORAGE) > 0:
            flag = False
        else:
            flag = True
        placeholder.set_visibility(flag)
        stored_memo_list.set_visibility(not flag)

    def delete_memo(title: str):
        result = _DO.delete_memo(USER_INS.uid, title)
        if result[0] == 'e':
            notification('negative', f'メモ削除にエラーが発生しました、処理を中止します。\n{result[1]}', multiline= True)
            return
        notification('positive', 'メモ削除しました。')
        MEMO_STORAGE.pop(title, '')
        refresh_memo()
        
        

    def refresh_memo():
        stored_memo_list.clear()
        active_placeholder()
        if len(MEMO_STORAGE) < 1:
            return
        #for k in MEMO_STORAGE.keys():
        for k, v in MEMO_STORAGE.items():
            with stored_memo_list:
                with ui.button_group().classes('max-w-56 h-8 truncate normal-case').props('push'):
                    with ui.button().classes('w-40 h-8 p-2').props('push') as new_button:
                        text_size = 'base'
                        display_title = v[1]
                        leng = len(display_title)
                        if leng > 19:
                            text_size = 'xs'
                        elif leng > 15:
                            text_size = 's'
                        elif leng > 9:
                            text_size = 'base'

                        ui.html(display_title, sanitize= False).classes(f'w-full h-8 text-{text_size} truncate normal-case inline-block align-left')
                        new_button.tooltip(display_title)
                    new_button.on_click(lambda t=k: retrive_memo(t))
                    close_btn = ui.button(icon='close', color='red', on_click=lambda t=k: delete_memo(t)).props('push')

    def analyse_memo(word_list:List[str], db_list:List[str]) -> str:
        holder = ''
        if len(word_list) > 0 and len(word_list) >= len(db_list):
            #print('reached 1')
            holder = f'{datetime.now().strftime('%Y-%m-%d')} {word_list[0]}'
            if len(word_list) > 1:
                holder += 'とその他'
        elif len(db_list) > 0 and len(word_list) == 0:
            holder = db_list[0]
            if len(db_list) > 1:
                holder += 'とその他'
        else:
            #print('reached 2')
            from wonderwords import RandomWord
            rand_word = RandomWord()
            dict_for_bp = {
                1 : 'u',
                2 : 'j',
                3 : 'r',
                4 : 'l',
                5 : 'b'
            }
            
            rand = rand_word.word()
            muse = dict_for_bp[random.randint(1, 5)]
            first = rand_word.word(starts_with= muse, include_parts_of_speech=['adjectives'])
            if muse == 'b':
                muse = 'p'
            last = rand_word.word(starts_with= muse, include_parts_of_speech=['nouns'])


            holder = f'メモ {first} {last}'
        
        return holder



    def save_memo():
        current_id = MEMO_OBJ.current_id
        if MEMO_OBJ is None or not MEMO_OBJ.value:
            return
        new_memo: str = MEMO_OBJ.value
        key_word_list = M.regex_split(new_memo, 'lot')
        key_db_list = M.regex_split(new_memo, 'prodcode')
        if current_id is None:
            title = analyse_memo(key_word_list, key_db_list)
            result = _DO.new_memo(USER_INS.uid, title, new_memo)
        else:
            title = MEMO_STORAGE[current_id][1]
            result = _DO.update_memo(USER_INS.uid, current_id, new_memo)
            MEMO_OBJ.current_id = None
        #print(f'*********************\nmemo contents = \n{new_memo}\n\ntitle = \n{title}\n\n*******************')
        
        if result[0] == 'e':
            notification('negative', f'メモ保存にエラーが発生しましたが、処理を続行します。\n{result[1]}', multiline= True)
            USER_INS.retry_log(_DO.new_memo, USER_INS.uid, title, new_memo)
        MEMO_STORAGE[result[1]] = (new_memo, title)
        MEMO_OBJ.set_value(None)
        refresh_memo()
    
    def retrive_memo(title: str):
        global MEMO_OBJ
        global MEMO_STORAGE

        current_memo = MEMO_OBJ.value
        
        if MEMO_OBJ is None:
            return
        retrived = MEMO_STORAGE.get(title, '')
        #print(retrived)
        if retrived == '':
            #need to implement a procedure to delete the empty note once found
            return
        
        if current_memo != '' and current_memo == retrived:
            #need to implement a dialog to warn user that the current memo will be overwritten 
            return

        else:
            #MEMO_OBJ.set_value(retrived)
            MEMO_OBJ.current_id = title
            MEMO_OBJ.value = retrived[0]
             
            #MEMO_OBJ.update()

    
    set_og_drawer(title, _list, False)
    with drawer:
        with ui.list().classes('h-full') as memo_list:
            memo_title = ui.label('メモ帳').classes('text-lg font-bold p-4')
            MEMO_OBJ = CustomEditor().classes('max-w-64').props('max-height=60vh')

            save_btn = ui.button('保存').classes('mt-2')
            close_btn = ui.button('閉じる').classes('mt-2').props('flat').on_click(close_memo)

            ui.separator().classes('w-full m-4')

            stored_memo_title = ui.label('メモ').classes('text-lg font-bold p-4')

            stored_memo_list = ui.scroll_area().classes('h-full')
            placeholder = ui.label('保存されたメモはありません。')
            
            refresh_memo()

            save_btn.on_click(save_memo)
    
#MARK: Supplier Schedule Drawer
async def supplier_sched_drawer(title: ui.label, _list: ui.list, drawer: ui.left_drawer):
    set_og_drawer(title, _list, False)
    with drawer:
        try:
            with ui.element('div').classes('p-4 w-full') as sched_panel:
                # ── Header + back button ────────────────────────────────────
                with ui.row().classes('items-center gap-2 mb-4'):
                    ui.button(icon='arrow_back', on_click=lambda: (
                        sched_panel.set_visibility(False),
                        set_og_drawer(title, _list, True)
                    )).props('flat round dense')
                    ui.label('仕入先スケジュール').classes('font-bold text-base')

                # ── Vendor section label ─────────────────────────────────────
                ui.label('仕入先').classes('text-xs text-gray-500')

                # ── Skeleton placeholders (shown while loading) ──────────────
                # gap-1 on the column is used instead of mb-1 on each skeleton —
                # equivalent visual spacing via flexbox gap.
                with ui.column().classes('w-full gap-1') as skeleton_area:
                    for _ in range(5):
                        ui.skeleton('rect').classes('w-full h-10')

                # ── Vendor button list (hidden until loaded) ─────────────────
                with ui.scroll_area().classes('w-full').style('max-height:60vh') as vendor_area:
                    vendor_col = ui.column().classes('w-full')
                vendor_area.set_visibility(False)

                # ── No-result / error label (hidden until needed) ────────────
                no_result_label = ui.label('結果なし').classes('text-sm text-gray-400 text-center mt-4')
                no_result_label.set_visibility(False)

        except Exception as e:
            notification('warning', f'予想外のエラーが発生しました：\n{e}')

    # ── Nested helpers ───────────────────────────────────────────────────────

    async def _on_vendor_click(vendor_code: str, vendor_name: str = ''):
        drawer.hide()
        spinner_open()
        sched_panel.set_visibility(False)
        set_og_drawer(title, _list, True)
        _today = date.today()
        params = {
            'vendor_code': vendor_code,
            'vendor_name': vendor_name,
            'date_from':   date(_today.year - 1, 1,  1).isoformat(),
            'date_to':     date(_today.year,     12, 31).isoformat(),
        }
        await _open_supplier_schedule_sample(params)

    async def _load_vendors():
        try:
            df = await M.fast_query('NOKEY_vendor_master', [], True)
            if df is not None and not df.empty:
                skeleton_area.set_visibility(False)
                for vendor_code, vendor_name in zip(df['VENDOR_CODE'], df['VENDOR_NAME']):
                    with vendor_col:
                        import asyncio
                        ui.button(
                            vendor_name,
                            on_click=lambda vc=vendor_code, vn=vendor_name: asyncio.ensure_future(_on_vendor_click(vc, vn))
                        ).classes('w-full').props('flat align=left')
                vendor_area.set_visibility(True)
            else:
                skeleton_area.set_visibility(False)
                no_result_label.set_visibility(True)
        except Exception:
            skeleton_area.set_visibility(False)
            no_result_label.set_visibility(True)

    await _load_vendors()


def tree_parser(tree_data:List[dict], tag: Literal['C', 'A', 'M', 'S']) -> str:
    dict_ = {
        'C' : '部品',
        'M' : '材料',
        'A' : '',
        'S' : '中間品'
    }
    output_text = ''
    key = dict_.get(tag, '')
    for node in tree_data:
        if key in node.get('type', '-'):
            output_text += f'{node.get('id', '-')}, {node.get('sap', '-')}, {node.get('usage', '-')}, {node.get('name', '-')}\n'
    return output_text


#MARK: Relation Drawer
def sap_relation_quick_drawer(title: ui.label, _list: ui.list, drawer: ui.left_drawer):
    def close_relation():
        relation_list.set_visibility(False)
        set_og_drawer(title, _list, True)

    def tree_event_receiver(event:events.GenericEventArguments, tree_obj: ui.tree):
        screen_element_copy(tree_parser(tree_obj.nodes(), event.args))

    
    async def relation_query():
        contents_area.clear()
        key = search_input.value
        try:
            key_obj = M.validate_text(key, 'p', '\'5501\', \'5401\'')
        except E.ExtraKeyExc:
            notification('warning', '一つのDB図番だけ入力してください。')
            skeletons_up(ske_list, False)
            raise
        except:
            return {}

    
        
        skeletons_up(ske_list, True)

        try:
            records = await uni_query(key_obj)
        except:
            skeletons_up(ske_list, False)
            return
        #print(records)

        data_rec = records['p']
        me:Tuple[str, str] = None
        top: List[str] = []
        bottom = False
        search_key = key_obj.key[0]
        holder_assy = []
        holder_compo = []
        holder_rawmate = []
        repeat = []

        #print(search_key)
        for item in data_rec:
            #print(item)
            #print(item[5], item[6])
            if me is None:
                if item[0] == search_key:
                    me = (item[0], item[1], item[5], item[6])
                if item[2] == search_key:
                    me = (item[2], item[3], item[5], item[6])
                #print('me defined')
            if item[2] == search_key:
                new_item = {'id': item[0],
                            'sap': item[1],
                            'usage': item[4],
                            'type': f'{"製品 - " if item[1].startswith("101") else "中間品 - " if item[1].startswith("102") else "部品 - " if item[1].startswith("103") else ""}',
                            'name': f'{item[5]} | {item[6]}'}
                top.append(new_item)
                #print('top added')
            if item[0] == search_key:
                new_item = {'id': item[5] if item[2] is None or item[2] == '*' else item[2],
                            'sap': item[3],
                            'usage': item[4],
                            'type': f'{"製品 - " if item[3].startswith("101") else "中間品 - " if item[3].startswith("102") else "部品 - " if item[3].startswith("103") else "材料/その他 - "}',
                            'name': f'{item[5]} | {item[6]}'}
                
                bottom = True
                sap_id: str = new_item['sap']
                if sap_id in repeat:
                    continue
                repeat.append(sap_id)
                if sap_id.startswith("102"):
                    holder_assy.append(new_item)
                if sap_id.startswith("103"):
                    holder_compo.append(new_item)
                if sap_id.startswith("104"):
                    holder_rawmate.append(new_item)

        skeletons_up(ske_list, False)

        with contents_area:
            product_header = ui.html(f'''
                <strong class= "text-blue-500 to-[#1CFCCC] text-xl font-extrabold">{me[0]}</strong>
                <p>{me[1]} : {'製品' if me[1].startswith('101') else '中間品' if me[1].startswith('102') else '部品' if me[1].startswith('103') else '材料/その他'}</p>
            ''', sanitize= False)
            product_header.on('dblclick', lambda: screen_element_copy(me[1]))
            if len(top) > 0:
                #holder = ''
                tree_node_used: Dict[str, str | List[str]] = [{'id': f'📦 {search_key}は、以下の製品に使用されます。'}]
                tree_node_used[0]['children'] = top
                
                used_in_tree = ui.tree(tree_node_used, label_key='id')
                used_in_tree.add_slot('default-header', '''
                    <span :props="props">{{props.node.type}} <strong>{{ props.node.id }}</strong>
                        <q-menu touch-position context-menu>
                            <q-item-label header>コピー</q-item-label>
                            <q-item v-if="!props.node.id.includes('🔧') && !props.node.id.includes('📦')" clickable v-close-popup @click="() => props.tree.$emit('copy', props.node.id)">
                                <q-item-section>{{props.node.id}}</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('部品')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'C')">
                                <q-item-section>部品ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('中間品')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'S')">
                                <q-item-section>中間品ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('材料')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'M')">
                                <q-item-section>材料ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('📦')" => props.tree.$emit('tcopy', 'A')">
                                <q-item-section>全部材ツリー</q-item-section>
                            </q-item>                                                                            
                        </q-menu>
                                     </span>
                    ''')
                used_in_tree.add_slot('default-body', '''
                    <span :props="props" v-if="!props.node.id.includes('品')">品目コード: "{{ props.node.sap }}" <br>数量: {{ props.node.usage }}
                        <q-menu touch-position context-menu>
                            <q-item-label header>コピー</q-item-label>
                            <q-item clickable v-close-popup @click="() => props.tree.$emit('copy', props.node.sap)">
                                <q-item-section>{{props.node.sap}}</q-item-section>
                            </q-item>
                        </q-menu>
                                      </span>
                    ''')

            #ui.label(f'{"製品 - " if me[1].startswith("101") else "中間品 - " if me[1].startswith("102") else "部品 - " if me[1].startswith("103") else "その他 - "} ||\n DB図番: {me[0]} |\n 品目コード {me[1]}')
            
            if bottom > 0:
                tree_node_b_used: Dict[str, str | List[str]] = [{'id': f'📦 {search_key}は、以下の中間品/部品を含んでいます。'}]
                tree_node_assy: Dict[str, str | List[str]] = {'id': '🔧 中間品'}
                tree_node_compo: Dict[str, str | List[str]] = {'id': '🔧 部品'}
                tree_node_rawmate: Dict[str, str | List[str]] = {'id': '🔧 材料'}
                tree_node_assy['children'] = holder_assy
                tree_node_compo['children'] = holder_compo
                tree_node_rawmate['children'] = holder_rawmate
                tree_node_b_used[0]['children'] = [tree_node_assy, tree_node_compo, tree_node_rawmate]
                    
                               
                b_used_tree = ui.tree(tree_node_b_used, label_key='id')
                b_used_tree.add_slot('default-header', '''
                    <span :props="props">{{props.node.type}} <strong>{{ props.node.id }}</strong>
                        <q-menu touch-position context-menu>
                            <q-item-label header>コピー</q-item-label>
                            <q-item v-if="!props.node.id.includes('🔧') && !props.node.id.includes('📦')" clickable v-close-popup @click="() => props.tree.$emit('copy', props.node.id)">
                                <q-item-section>{{props.node.id}}</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('部品')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'C')">
                                <q-item-section>部品ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('中間品')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'S')">
                                <q-item-section>中間品ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('材料')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'M')">
                                <q-item-section>材料ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('📦')" => props.tree.$emit('tcopy', 'A')">
                                <q-item-section>全部材ツリー</q-item-section>
                            </q-item>                                                                            
                        </q-menu>
                                     </span>
                    ''')
                b_used_tree.add_slot('default-body', '''
                    <span :props="props" v-if="!props.node.id.includes('🔧') && !props.node.id.includes('📦')">品目コード: {{ props.node.sap }} <br>名称: {{ props.node.name }}<br>数量: {{ props.node.usage }}
                        <q-menu touch-position context-menu>
                            <q-item-label header>コピー</q-item-label>
                            <q-item v-if="!props.node.id.includes('🔧') && !props.node.id.includes('📦')" clickable v-close-popup @click="() => props.tree.$emit('copy', props.node.id)">
                                <q-item-section>{{props.node.id}}</q-item-section>
                            </q-item>
                            <q-item v-if="!props.node.id.includes('🔧') && !props.node.id.includes('📦')" clickable v-close-popup @click="() => props.tree.$emit('copy', props.node.sap)">
                                <q-item-section>{{props.node.sap}}</q-item-section>
                            </q-item>                                       
                            <q-item v-if="!props.node.id.includes('🔧') && !props.node.id.includes('📦')" clickable v-close-popup @click="() => props.tree.$emit('copy', props.node.name)">
                                <q-item-section>{{props.node.name}}</q-item-section>
                            </q-item>                                     
                            <q-item v-else-if="props.node.id.includes('部品')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'C')">
                                <q-item-section>部品ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('中間品')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'S')">
                                <q-item-section>中間品ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('材料')" clickable v-close-popup @click="() => props.tree.$emit('tcopy', 'M')">
                                <q-item-section>材料ツリー</q-item-section>
                            </q-item>
                            <q-item v-else-if="props.node.id.includes('📦')" => props.tree.$emit('tcopy', 'A')">
                                <q-item-section>全部材ツリー</q-item-section>
                            </q-item>                                                                            
                        </q-menu>
                    </span>
                    
                    ''')
                b_used_tree.on('copy', lambda e:screen_element_copy(e.args))
                b_used_tree.on('tcopy', lambda e:tree_event_receiver(e, b_used_tree))
                #b_used_tree.add_slot('default-header', '''
                #    <span :props="props">{{props.node.type}} <strong>{{ props.node.id }}</strong></span>
                #''')

                #b_used_tree.add_slot('default-body', '''
                #    <span :props="props" v-if="!props.node.id.includes('🔧')" @contextmenu.prevent="() => props.tree.$emit('show_menu', props.node)">
                #        品目コード: "{{ props.node.sap }}" <br>名称: {{ props.node.name }}<br>数量: {{ props.node.usage }}
                #    </span>
                #''')
                #print(tree_node_b_used)
                
                
    set_og_drawer(title, _list, False)
    with drawer:
        with ui.list() as relation_list:
            relation_title = ui.label('品番/品目関係検索').classes('text-lg font-bold p-4')
            with ui.row().classes('w-full'):
                search_input = ui.input(placeholder='DB図番を検索', validation=
                                        lambda value: 'DB図番ではありません' if len(M.regex_split(value, 'prodcode')) == 0  else '1つのDB図番だけ入力してください' if len(M.regex_split(value, 'prodcode')) > 1 else None
                                        ).classes('w-full').props('rounded outlined dense').style('min-width: 250px')
                
            with ui.row().classes('w-full'):
                search_btn = ui.button('検索').props('rounded outlined color=primary')
                ui.space()
                close_btn = ui.button('閉じる', color='red', on_click=close_relation).props('rounded outlined')

            ui.separator().classes('w-full m-4')

            with ui.element('div').classes('w-full') as display_container:
                skeleton1 = ui.skeleton('text').classes('w-full')
                skeleton2 = ui.skeleton('text').classes('w-full')
                skeleton3 = ui.skeleton('text').classes('w-full')
                ske_list: List[ui.skeleton] = [skeleton1, skeleton2, skeleton3]
                skeletons_up(ske_list, False)
                contents_area = ui.element('div').classes('w-full')
                

            search_btn.on_click(relation_query)

#MARK: Flow and process drawer
def flow_and_process_code():
    class Product():
        def __init__(self, data_tuple: Tuple[str]):
            #print(data_tuple)
            self.header_SAP = data_tuple[0]
            self.header_DB = data_tuple[1]
            self.name = data_tuple[2]
            self.flow: Dict[str, Flow] = {data_tuple[3]: Flow(data_tuple)}
        
        def add_flow(self, data_tuple: Tuple[str]):
            #print(f'Adding flow with code {data_tuple[3]} to product {self.name}')
            newflowins = self.flow.get(data_tuple[3], Flow(data_tuple))
            newflowins.add_process(data_tuple)
            self.flow[data_tuple[3]] = newflowins
            #print(f'Current flows for product {self.name}: {list(self.flow.keys())}')
        
        def output(self) -> Dict[str, dict]:
            print(f'Outputting product {self.name} with SAP {self.header_SAP} and DB {self.header_DB}')
            root = {
            'id': '0',
            'db': self.header_DB,
            'sap' : self.header_SAP,
            'name' : self.name,
            'flowcode': 'NA',
            'flowname' : 'NA',
            'manu_type' : 'NA',
            'processcode' : 'NA',
            'equipment' : 'NA',
            'labortime' : 'NA',
            'machinetime' : 'NA',
            'totaltime' : 'NA',
            'children' : []
            }

            for flow_idx, flow_ins in enumerate(self.flow.values()):
                newflow = {
                'id': f'1-{flow_idx}',
                'db': 'NA',
                'sap' : 'NA',
                'name' : 'NA',
                'flowcode': flow_ins.flowcode,
                'flowname' : flow_ins.flowname,
                'manu_type' : flow_ins.flowtype,
                'processcode' : 'NA',
                'equipment' : 'NA',
                'labortime' : 'NA',
                'machinetime' : 'NA',
                'totaltime' : 0,
                'children' : []
                }
                root['children'].append(newflow)
                for proc_idx, process_ in enumerate(flow_ins.process.values()):
                    try:
                        newprocess = {
                        'id': f'2-{flow_idx}-{proc_idx}',
                        'db': 'NA',
                        'sap' : 'NA',
                        'name' : 'NA',
                        'flowcode': 'NA',
                        'flowname' : 'NA',
                        'manu_type' : 'NA',
                        'processcode' : process_.processcode,
                        'processname' : process_.processname,
                        'equipment' : 'NA',
                        'labortime' : f'{process_.time[0] or 0:.1f}分',
                        'machinetime' : f'{process_.time[1] or 0:.1f}分',
                        'totaltime' : f'{process_.time[2] or 0:.1f}分',
                        'children' : []
                        }
                        newflow['totaltime'] += float(process_.time[2] or 0)
                        newflow['children'].append(newprocess)
                        for equip_idx, equipment_ in enumerate(process_.equipment):
                            newequipment = {
                            'id': f'3-{flow_idx}-{proc_idx}-{equip_idx}',
                            'db': 'NA',
                            'sap' : 'NA',
                            'name' : 'NA',
                            'flowcode': 'NA',
                            'flowname' : 'NA',
                            'manu_type' : 'NA',
                            'processcode' : 'NA',
                            'equipment' : equipment_,
                            'labortime' : 'NA',
                            'machinetime' : 'NA',
                            'totaltime' : 'NA',
                            'children' : []
                            }
                            newprocess['children'].append(newequipment)
                    except Exception as e:
                        print(f'Error processing flow {flow_ins.flowcode} process {process_.processcode}: {e}')
            print(f'Finished outputting product {self.name}')
            return root

    class Flow():
        def __init__(self, data_tuple: Tuple[str]):
            self.flowcode = data_tuple[3]
            self.flowname = data_tuple[4]
            self.flowtype = data_tuple[5]
            self.process: Dict[str, Process] = {data_tuple[7]: Process(data_tuple)}
        
        def add_process(self, data_tuple: Tuple[str]):
            newprocessins = self.process.get(data_tuple[7], Process(data_tuple))
            newprocessins.add_equipment(data_tuple)
            self.process[data_tuple[7]] = newprocessins

    class Process():
        def __init__(self, data_tuple: Tuple[str]):
            self.processcode = data_tuple[7]
            self.SEQ = data_tuple[6]
            self.processname = data_tuple[8]
            self.time = (data_tuple[10], data_tuple[11], data_tuple[12])
            self.equipment: List[str] = [data_tuple[9]]
        def add_equipment(self, data_tuple: Tuple[str]):
            if not data_tuple[7] == self.processcode:
                raise Exception(f'Process code doesn\'t match \n Instance process code = {self.processcode} \n Input process code = {data_tuple[7]}')
            if data_tuple[9] not in self.equipment:
                self.equipment.append(data_tuple[9])

    async def flow_query(container: ui.card, skeletons:List[ui.skeleton]) -> List[Dict[str, str | dict]]:
        container.clear()
        key = search_box.value
        plant_str = plant.value
        if not plant_str:
            notification('warning', '工場区分を選択してください。')
            return 'NoPlant'
        try:
            key_obj = M.validate_text(key, 'f', plant_str)
        except E.ExtraKeyExc:
            notification('warning', '一つのDB図番だけ入力してください。')

        except:
            return {}
    
        
        skeletons_up(skeletons, True)

        try:
            records = await uni_query(key_obj)
        except:
            skeletons_up(skeletons, False)
            return

        root = None
        for data in records['f']:
            if root is None:
                root = Product(data)
            root.add_flow(data)
        
        asdict = root.output()
        
        with container:
            with ui.row().classes('items-center gap-2 q-mt-sm'):
                flowtree_select_all = ui.chip(icon='checklist',     color='orange', text='全選択')
                flowtree_deselect_all = ui.chip(icon='cancel',        color='grey',   text='全取消')
                flowtree_copy = ui.chip(icon='content_copy',  color='green',  text='コピー')
                flowtree_csv = ui.chip(icon='file_download', color='blue', text='CSV出力')

            
            flowtree = ui.tree([asdict], label_key='id', node_key='id', tick_strategy='leaf')
            flowtree.add_slot('default-header', '''
                    <span :props="props" v-if="props.node.id == '0'">{{ props.node.db }}<br>{{ props.node.sap }}<br>📦 <strong>{{ props.node.name }}</strong></span>
                    <span :props="props" v-if="props.node.id.startsWith('1')">{{ props.node.flowcode }}<br>{{ props.node.manu_type }}<br>🧾️ <strong>{{ props.node.flowname }}</strong></span>
                    <span :props="props" v-if="props.node.id.startsWith('2')">{{ props.node.processcode }}<br>🛠️ <strong>{{ props.node.processname }}</strong><br>作業時間：手作業　{{ props.node.labortime }} ｜　機械作業　{{ props.node.machinetime }}｜　合計時間　{{ props.node.totaltime }}</span>
                    <span :props="props" v-if="props.node.id.startsWith('3')">使用設備：<br>⚙️ <strong>{{ props.node.equipment }}</strong></span>
                                        ''')
            flowtree.add_slot('default-body', '''
                    <span :props="props" v-if="props.node.id.startsWith('1')">フロー合計作業時間: "{{ props.node.totaltime.toFixed(2) }}"分</span>
                    ''')
            print('tree rendered')
            async def _copy_flowtree():
                ticked = await flowtree.run_method('getTickedNodes')
                if not ticked:
                    notification('negative', '選択中の階層はありません。')
                    return
                ticked_ids = {node['id'] for node in ticked}
                lines = []
                for flow_node in asdict.get('children', []):
                    flow_header_added = False
                    for proc_node in flow_node.get('children', []):
                        proc_header_added = False
                        for equip_node in proc_node.get('children', []):
                            if equip_node['id'] in ticked_ids:
                                if not flow_header_added:
                                    lines.append(f"[フロー] {flow_node['flowcode']}: {flow_node['flowname']} ({flow_node['manu_type']})")
                                    flow_header_added = True
                                if not proc_header_added:
                                    lines.append(
                                        f"  [工程] {proc_node['processcode']}: {proc_node['processname']} "
                                        f"| 手作業: {proc_node['labortime']} | 機械: {proc_node['machinetime']} | 合計: {proc_node['totaltime']}"
                                    )
                                    proc_header_added = True
                                lines.append(f"    [設備] {equip_node['equipment']}")
                screen_element_copy('\n'.join(lines))

            async def _export_flowtree_csv():
                ticked = await flowtree.run_method('getTickedNodes')
                if not ticked:
                    notification('negative', '選択中の階層はありません。')
                    return
                ticked_ids = {node['id'] for node in ticked}
                rows = []
                for flow_node in asdict.get('children', []):
                    for proc_node in flow_node.get('children', []):
                        for equip_node in proc_node.get('children', []):
                            if equip_node['id'] in ticked_ids:
                                rows.append({
                                    'フローコード': flow_node.get('flowcode', ''),
                                    'フロー名': flow_node.get('flowname', ''),
                                    '製造タイプ': flow_node.get('manu_type', ''),
                                    '工程コード': proc_node.get('processcode', ''),
                                    '工程名': proc_node.get('processname', ''),
                                    '手作業時間': proc_node.get('labortime', ''),
                                    '機械時間': proc_node.get('machinetime', ''),
                                    '合計時間': proc_node.get('totaltime', ''),
                                    '設備': equip_node.get('equipment', ''),
                                })
                if rows:
                    export_csv(pd.DataFrame(rows), 'flow')
                else:
                    notification('warning', 'CSVに出力するデータがありません。')

            flowtree_select_all.on_click(flowtree.tick)
            flowtree_deselect_all.on_click(flowtree.untick)
            flowtree_copy.on_click(_copy_flowtree)
            flowtree_csv.on_click(_export_flowtree_csv)
        skeletons_up(skeletons, False)
        
        
        


    with ui.dialog().classes('w-full h-5/6').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:
        close_btn = ui.label('閉じる').classes('text-xl text-white/75 hover:text-white cursor-pointer')
        close_btn.on('click', dialog.close)
        with ui.card().classes('w-full h-5/6').style('max-width: none'):
            ui.label('フロー/工程検索')
            ui.separator()
            with ui.row():
                search_box = ui.input('DB図番を入力してください', 
                                      validation=lambda value: 'DB図番ではありません' if len(M.regex_split(value, 'prodcode')) == 0  else '1つのDB図番だけ入力してください' if len(M.regex_split(value, 'prodcode')) > 1 else None
                                        ).classes('w-1/2').props('rounded outlined dense').style('min-width: 250px')
                excute_btn = ui.button('検索')
                plant = ui.toggle({'\'5501\'': '伊勢原', '\'5401\'': '宮田'})
            with ui.card().classes('w-full h-full shadow-xl') as container_c:
                skeleton1 = ui.skeleton('text').classes('w-64 h-38')
                skeleton2 = ui.skeleton('text').classes('w-64 h-38')
                skeleton3 = ui.skeleton('text').classes('w-64 h-32')
                skeleton4 = ui.skeleton('text').classes('w-32 h-32')
                skeleton5 = ui.skeleton('text').classes('w-64 h-32')
                container = ui.scroll_area().classes('w-full h-full')  
                ske_list: List[ui.skeleton] = [skeleton1, skeleton2, skeleton3, skeleton4, skeleton5]
                skeletons_up(ske_list, False)
            excute_btn.on_click(lambda c= container, sk= ske_list: flow_query(c, sk))
    dialog.open()


#MARK: Data viewer
class DataViewer():
    __slots__ = [
    'container',
    'current_ag',
    'ag_object',
    'ag_up',
    'ag_container',
    'current_flag',
    'splitter',
    'splitter_container',
    'refer_to_ag',
    'refer_to_agcon',
    'csv_button',
    'dataframe',
    'condition_exp',
    'setting_container'
    ]
    def __init__(self):
        self.container: ui.dialog = None
        self.current_ag: ui.aggrid = None
        self.ag_object: AgGrid = None
        self.ag_up: bool = False
        self.ag_container: ui.row = None 
        self.current_flag: str = None
        self.splitter: ui.splitter = None
        self.splitter_container: ui.card = None
        self.refer_to_ag: ui.aggrid = None
        self.refer_to_agcon: ui.element = None
        self.csv_button: ui.button = None
        self.dataframe: pd.DataFrame = None

    def init_blank(self):
        with ui.dialog().classes('w-full h-5/6').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:
            close_btn = ui.label('閉じる').classes('text-xl text-white/75 hover:text-white cursor-pointer')
            close_btn.on('click', dialog.close)
            with ui.card().classes('w-full h-5/6').style('max-width: none'):
                ui.label('データビュアー').bind_visibility_from(self, 'ag_up', value=False)
                ui.separator().bind_visibility_from(self, 'ag_up', value=False)
                with ui.row():
                    ui.button('黄袋システム', on_click=lambda: self.query_ag('D'))
                    ui.button('Kシステム', on_click=lambda: self.query_ag('K'))
                    ui.button('除湿室データ', on_click=lambda: self.query_ag('F', '指図'))
                    ui.button('運搬一覧').set_enabled(False)
                    ui.button('作業計画', on_click=lambda currentdate = datetime.now().strftime('%Y-%m-%d'): self.query_ag('S', search_condition=[currentdate]))
                    ui.separator().props('vertical')
                    self.csv_button = ui.button('CSVに出力', color='orange', on_click= lambda click_event: export_csv(self.dataframe, 'データビュアー出力'))

                with ui.expansion('オプション').classes('w-full').bind_visibility_from(self, 'ag_up', value=True) as self.condition_exp:
                    with ui.card().classes('w-full'):
                        self.setting_container = ui.card().classes('w-full')

                container_data_viewer = ui.card().classes('w-full h-full shadow-xl').bind_visibility_from(self, 'ag_up', value=True)
                    
                ui.label('リストから閲覧するデータを選択してください。').classes('w-full h-full text-5xl font-semibold place-content-center mb-6 pb-3 text-gray-300').bind_visibility_from(self, 'ag_up', value=False)

        self.container = container_data_viewer
        dialog.on('before-hide', self.deactivate_splitter)
        self.csv_button.disable()
        dialog.open()
    

    def deactivate_splitter(self):
        if self.splitter:
            self.ag_container.move(self.container)
            self.transfer_ag(direction='push')
            #self.splitter_container.update()
            #self.splitter_container.clear()
            self.splitter_container.delete()
            self.splitter = None
            
            
    def activate_splitter(self):

        if len(current_tabs) == 1:
            notification('negative', '分割ビューできる検索中のデータはありません、先にデータを検索してください。')
            return
        with self.container:
            with ui.card().classes('w-full h-full shadow-xl') as self.splitter_container:
                with ui.splitter().classes('w-full') as self.splitter:
                    with self.splitter.before:
                        ui.label('データビュアー').classes('text-lg')
                        before_card = ui.card().classes('w-full h-full')
                        self.ag_container.move(before_card)
                    with self.splitter.after:
                        ui.label('　検索データ').classes('text-lg')
                        with ui.row().classes('w-full h-full') as tab_list:
                            with ui.card().classes('w-full h-full') as after_card:
                                with ui.element('div') as button_holder:
                                    for k, tb in current_tabs.items():
                                        if k == 'default_0':
                                            continue
                                        tab_name = tb[0][1]
                                        if tab_name == '指図関係':
                                            continue
                                        tabbt = ui.label(tab_name).classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg').on('click', 
                                                lambda tab_pnl = tb[2], to_container = after_card, bt_container=button_holder, direction = 'pull': self.transfer_ag(tab_pnl, to_container, bt_container, direction))
    
    def transfer_ag(self, tab_pnl: CustomTabPanel = None, transfer_to_con: ui.card = None, button_container: ui.element = None, direction: Literal['push', 'pull'] = 'pull'):

        if direction == 'pull':
            if button_container is None or tab_pnl is None:
                return
            target_ag: ui.aggrid = tab_pnl.aggrid_instance
            target_container = tab_pnl.aggrid_container
            button_container.delete()
            self.refer_to_ag = target_ag
            self.refer_to_agcon = target_container
            with transfer_to_con:
                recreate_con = ui.element('div').classes('w-full h-full flex-grow overflow-hidden')
                target_ag.move(recreate_con)
        if direction == 'push':
            if self.refer_to_ag is None or self.refer_to_agcon is None:
                return
            self.refer_to_ag.move(self.refer_to_agcon)
            self.refer_to_ag = None
            self.refer_to_agcon = None

                    
    def split_view_button(self, splitviewbt: ui.label):
        if self.splitter:
            splitviewbt.set_text('検索中のデータ参照')
            self.deactivate_splitter()
        
        else:
            splitviewbt.set_text('分割ビューを終了')
            self.activate_splitter()


    def setup_setting_container(self, key: str, s_viewer_date: str = None, filter_group: pd.DataFrame = None):
        self.setting_container.clear()
        if key == 'S':
            self.condition_exp.enable()
            self.condition_exp.text = '作業計画'
            with self.setting_container:
                view_date_container, date_picker = datepicker_input('日付', s_viewer_date)
                with view_date_container:
                    ui.icon('search').classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg').on(
                        'click', lambda new_date = date_picker.text[0]: self.query_ag('S', search_condition=[date_picker.text])
                    )
                ui.separator().classes('w-full')
                ui.label('工程グループフィルター').classes('text-base')
                with ui.row():
                    for gr in filter_group:
                        filter_bt = ui.label(str(gr)).classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg')
                        filter_bt.on('click', lambda condition = str(gr): self.ag_object.filter_api('大区分名称', 
                                                                                {'filterType': 'text', 'type': 'equals', 'filter': condition}))
                        
                ui.separator()
        
        if key == 'K':
            self.condition_exp.enable()
            self.condition_exp.text = 'Kシステム'
            with self.setting_container:
                ui.label('客先フィルター').classes('text-base')
                with ui.row():
                    for gr in filter_group:
                        filter_bt = ui.label(str(gr)).classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg')
                        filter_bt.on('click', lambda condition = str(gr): self.ag_object.filter_api('CUSTOMERS_DESCRIPTION', 
                                                                                {'filterType': 'text', 'type': 'equals', 'filter': condition}))
                        
            ui.separator()
                        
        if key == 'D':
            self.condition_exp.enable()
            self.condition_exp.text = '黄袋システム'
            with self.setting_container:
                ui.label('ファストフィルター').classes('text-base')
                with ui.row():
                    restricted = ui.label('出荷制限中').classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg').on('click', 
                                    lambda condition = '0': self.ag_object.filter_api('RELEASE_FLG', 
                                    {'filterType': 'text', 'type': 'equals', 'filter': condition}))
                    released = ui.label('制限解除済').classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg').on('click', 
                                    lambda condition = '1': self.ag_object.filter_api('RELEASE_FLG', 
                                    {'filterType': 'text', 'type': 'equals', 'filter': condition}))
            ui.separator()
        
        if key == 'F':
            self.condition_exp.text = '除湿室データ'
            #self.condition_exp.disable()

        with self.setting_container:
            splitview = ui.label('検索中のデータ参照' if self.splitter is None else '分割ビューを終了').classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg')
            splitview.on('click', lambda s = splitview: self.split_view_button(s))
            
            
                        
        
    async def query_ag(self, key: str, _row_id: str = 'index_id', search_condition: str = ''):
        result = await query_whole_table_wo_vali(key, _key = search_condition)
        self.deactivate_splitter()
        ag = AgGrid()
        filter_group = None
        if self.current_ag is not None:
            self.container.clear()
            self.current_ag = None
            self.dataframe = None
            #self.current_ag.delete()
            #self.setting_container.update()


        if key == 'S':
            filter_group = result['大区分名称'].unique()
        
        if key == 'K':
            filter_group = result['CUSTOMERS_DESCRIPTION'].unique()

        with self.container:
            self.ag_up = True
            ag.init_column_settings(key, True)
            grid_container, aggrid = ag.aggrid_creator(result, _table_view = True, row_id=_row_id)
            ag.cellkey_func_bind()
            self.ag_container = grid_container
            self.current_ag = aggrid
            self.ag_object = ag
            self.setup_setting_container(key, search_condition, filter_group)
        self.dataframe = result
        self.current_flag = key
        self.csv_button.enable()
        #aggrid.on('cellClicked', lambda e: print(f"Ctrl: {e.args.get('event', {}).get('ctrlKey')}, Shift: {e.args.get('event', {}).get('shiftKey')} event = {e}"))
    



def dataviewer():
    try:
        newdataviewer = DataViewer()
        newdataviewer.init_blank()
    except Exception as e:
        notification('warning', f'データビュアーを立ち上げる時に以下のエラーが発生しました：{e}')

#MARK: Settings Drawer
def setting_screen() -> ui.dialog:
    schema_ = {
    "type": "object",
        "properties": {
            "マーク色スキーム": {
                "type": "object",
                "patternProperties": {
                    "^[^<>\\/\\.\\[\\]{}()@#$%^&*+=|\\\\:;\"'`~,?!]+$": {
                    "type": "string",
                    "pattern": "^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$",
                    "description": "Hex color code in format #RGB or #RRGGBB"
                    }
                },
                "additionalProperties": False,
                "minProperties": 1
                },
                "仕掛検索デフォルト": {
                    "type": "object",
                    "properties": {
                        "ise": {
                            "type": "boolean"
                        },
                        "miya": {
                            "type": "boolean"
                        },
                        "prestart": {
                            "type": "boolean"
                        }
                },
                "additionalProperties": False,
                "minProperties": 3
            },
            
            },
            "required": ["マーク色スキーム", "仕掛検索デフォルト"],
            "additionalProperties": False
            }
    
    help_string = {'content': '- 任意の項目をクリックして説明を表示します。'}

    def help_string_generator(event: events.JsonEditorSelectEventArguments):
        try:
            path = event.selection
            new_md = ''
            if 'path' in path:
                #print(len(path), path['path'])
                if len(path) < 1 or not path['path']:
                    #print(1)
                    new_md = '- 任意の項目をクリックして説明を表示します。'
                else:
                    #print(2)
                    #print(path['path'][0] == 'マーク色スキーム')
                    if path['path'][0] == 'マーク色スキーム':
                        #print(3)
                        new_md = '- マーク色の変更方法<br> - マークの名前：マークのセットでマーク色を変更できます。<br> - マーク色：色をHEXコードで入力してください、既存の色を変更する場合は左の色サンプルをクリックして選択することができます。'
                        
                    if path['path'][0] == '仕掛検索デフォルト':
                        new_md = '- 仕掛検索デフォルト検索の変更方法<br> - ise: 伊勢原工場, miya: 宮田工場：　左のチェックボックスにチェックを入れると、仕掛検索画面ではデフォルトで有効/無効な状態となります。<br> - prestart: チェックを入れるとデフォルトでは投入前の指図表示のオプションが有効になります。'
                    if 'edit' in path:
                        new_md += '<br>- 名前以外の直接の値変更は**お勧めではありません**、できるだけカラーピッカーか、チェックボックスを使用してください。<br>直接入力する場合は、**JSONのSyntaxに従って、真偽値を小文字で入力してください。**'
            else:
                new_md = '- Treeモードから離脱したため、説明分の表示が無効となります。<br>編集する際は**JSON**に従ってください。<br>オプション式の編集を実施するには、Treeモードに切り替えしてください。'
            
            help_string['content'] = new_md
        except Exception as e:
            notification('negative', f'処理中にエラーが発生しました：{e}')
           

        
    with ui.dialog().classes('w-full').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:
        close_btn = ui.label('閉じる').classes('text-xl text-white/75 hover:text-white cursor-pointer')
        close_btn.on('click', dialog.close)
        with ui.card().classes('w-full h-5/6').style('max-width: none'):
            with ui.card().classes('w-full'):
                ui.label('設定').classes('text-lg font-semibold')
            with ui.tabs() as settings_tabs:
                tab_general = ui.tab('一般設定')
                tab_filter = ui.tab('フィルターメモリ')
            with ui.tab_panels(settings_tabs, value=tab_general).classes('w-full flex-1'):
                with ui.tab_panel(tab_general):
                    with ui.card().classes('w-full flex'), ui.row().classes('w-full flex'):
                        options_for_edit = {
                                'マーク色スキーム': dict(TAG_COLOR),
                                '仕掛検索デフォルト': dict(wip_option)}
                        editor = ui.json_editor({
                            'content': {'json': options_for_edit}
                        },schema= schema_).classes('w-1/2 flex-1')
                        with ui.card().classes('w-1/2 flex-1'), ui.row():
                            save = ui.button('保存')
                            newcolor = ui.button('新規マークカラーを追加')
                            ui.separator().classes('w-full')
                            ui.markdown().bind_content(help_string, strict=False).classes('w-full')
                with ui.tab_panel(tab_filter):
                    with ui.row().classes('w-full items-center mb-2'):
                        ui.label('保存済みフィルター').classes('text-lg font-semibold flex-1')
                        refresh_btn = ui.button('更新', icon='refresh')
                    with ui.scroll_area().style('width:100%;height:400px;'):
                        with ui.column().classes('w-full') as filter_list_container:
                            pass

    editor.on_select(help_string_generator)

    def _create_filter_row(f):
        mode_labels = {'i': '指図', 'w': '仕掛'}
        mode_label = mode_labels.get(f.query_mode, f.query_mode)
        scope_text = 'グローバル' if f.db_scope is None else f.db_scope
        with ui.card().classes('w-full mb-2'):
            with ui.row().classes('w-full items-center gap-2'):
                ui.label(f.filter_name).classes('font-medium flex-1')
                ui.label(f'モード: {mode_label}').classes('text-sm text-gray-500 px-2')
                ui.label(f'スコープ: {scope_text}').classes('text-sm text-gray-500 px-2')

                async def rename(fo=f):
                    with ui.dialog() as rename_dialog, ui.card():
                        ui.label('フィルター名変更').classes('text-lg font-semibold')
                        new_name_input = ui.input('新しい名前', value=fo.filter_name)
                        async def do_rename(fobj=fo):
                            if not new_name_input.value:
                                ui.notify('フィルター名を入力してください', type='warning')
                                return
                            result = _DO.update_filter_memory(USER_INS.uid, fobj.filter_id, filter_name=new_name_input.value)
                            if result[0] == 's':
                                ui.notify('名前を変更しました', type='positive')
                                rename_dialog.close()
                                refresh_filter_table()
                            else:
                                ui.notify(f'エラー: {result[1]}', type='negative')
                        with ui.row():
                            ui.button('保存', on_click=do_rename)
                            ui.button('キャンセル', on_click=rename_dialog.close).props('flat')
                    rename_dialog.open()

                async def delete_confirm(fo=f):
                    with ui.dialog() as del_dialog, ui.card():
                        ui.label(f'「{fo.filter_name}」を削除しますか？').classes('text-lg')
                        with ui.row():
                            async def do_delete(fobj=fo):
                                result = _DO.delete_filter_memory(USER_INS.uid, fobj.filter_id)
                                if result[0] == 's':
                                    ui.notify('フィルターを削除しました', type='positive')
                                    del_dialog.close()
                                    refresh_filter_table()
                                else:
                                    ui.notify(f'エラー: {result[1]}', type='negative')
                            ui.button('削除', on_click=do_delete, color='red')
                            ui.button('キャンセル', on_click=del_dialog.close).props('flat')
                    del_dialog.open()

                async def edit_condition(fo=f):
                    with ui.dialog() as edit_dialog, ui.card():
                        ui.label('フィルター条件を編集').classes('text-lg font-semibold')
                        filter_schema = {
                            "type": "object",
                            "additionalProperties": {
                                "oneOf": [
                                    {
                                        "type": "object",
                                        "required": ["filterType", "type"],
                                        "properties": {
                                            "filterType": {
                                                "type": "string",
                                                "enum": ["text", "number", "date"]
                                            },
                                            "type": {
                                                "type": "string",
                                                "enum": [
                                                    "contains", "notContains", "equals", "notEqual",
                                                    "startsWith", "endsWith", "blank", "notBlank",
                                                    "lessThan", "lessThanOrEqual",
                                                    "greaterThan", "greaterThanOrEqual", "inRange"
                                                ]
                                            },
                                            "filter": {},
                                            "filterTo": {}
                                        }
                                    },
                                    {
                                        "type": "object",
                                        "required": ["filterType", "operator", "conditions"],
                                        "properties": {
                                            "filterType": {
                                                "type": "string",
                                                "enum": ["text", "number", "date"]
                                            },
                                            "operator": {
                                                "type": "string",
                                                "enum": ["AND", "OR"]
                                            },
                                            "conditions": {
                                                "type": "array",
                                                "minItems": 2,
                                                "items": {
                                                    "type": "object",
                                                    "required": ["filterType", "type"],
                                                    "properties": {
                                                        "filterType": {
                                                            "type": "string",
                                                            "enum": ["text", "number", "date"]
                                                        },
                                                        "type": {
                                                            "type": "string",
                                                            "enum": [
                                                                "contains", "notContains", "equals", "notEqual",
                                                                "startsWith", "endsWith", "blank", "notBlank",
                                                                "lessThan", "lessThanOrEqual",
                                                                "greaterThan", "greaterThanOrEqual", "inRange"
                                                            ]
                                                        },
                                                        "filter": {},
                                                        "filterTo": {}
                                                    }
                                                }
                                            }
                                        }
                                    }
                                ]
                            }
                        }
                        editor = ui.json_editor({
                            'content': {'json': json.loads(fo.filter_model)}
                        }, schema=filter_schema).classes('w-full')
                        async def do_save_condition(fobj=fo):
                            data = await editor.run_editor_method('get')
                            vali = await editor.run_editor_method('validate')
                            if vali is not None:
                                if vali.get('contentErrors'):
                                    ui.notify('JSONが不正です', type='negative')
                                    return
                                if vali.get('validationErrors'):
                                    ui.notify('フィルター条件の形式が不正です。', type='warning')
                                    ui.notify(
                                        'filterType (text/number/date), type または operator+conditions の構成を確認してください。',
                                        type='info'
                                    )
                                    return
                            new_json = json.dumps(data['json'])
                            result = _DO.update_filter_memory(USER_INS.uid, fobj.filter_id, filter_model=new_json)
                            if result[0] == 's':
                                ui.notify('フィルター条件を更新しました', type='positive')
                                edit_dialog.close()
                                refresh_filter_table()
                            else:
                                ui.notify(f'エラー: {result[1]}', type='negative')
                        with ui.row():
                            ui.button('保存', on_click=do_save_condition)
                            ui.button('キャンセル', on_click=edit_dialog.close).props('flat')
                    edit_dialog.open()

                ui.button('名前変更', icon='edit', on_click=rename)
                ui.button('削除', icon='delete', on_click=delete_confirm, color='red')
                ui.button('条件編集', icon='tune', on_click=edit_condition)

    def refresh_filter_table():
        filter_list_container.clear()
        filters = _DO.get_filter_memories(USER_INS.uid)
        with filter_list_container:
            if not filters:
                ui.label('保存済みフィルターはありません。').classes('text-gray-500 p-4')
            else:
                for f in filters:
                    _create_filter_row(f)

    refresh_btn.on_click(refresh_filter_table)
    refresh_filter_table()

    async def save_setting():
        global TAG_COLOR
        global wip_option
        data = await editor.run_editor_method('get')
        # Current Data(After editing)
        inside_structure: Dict[str, Dict[str, str]] = data['json']
        # Previous Data(Editor be initialised with)
        options: Dict[str, Dict[str, str]] = dict(options_for_edit)
        # Pop the placeholder data from current data
        inside_structure['マーク色スキーム'].pop('//新規（名前を変更してください', '')
        # Refill the editor with altered(without placeholder)
        editor.properties['content']['json'] = inside_structure
        #editor.update()
        # Revalidate the data to check if schema is met
        vali = await editor.run_editor_method('validate')
        # If not met
        if vali is not None and 'validationErrors' in vali:
            notification('warning', '変更しようとしているデータは不正です。')
            notification('info', '不正データが含まれるため、前のバージョンに回復します。')
            editor.properties['content']['json'] = options
            #editor.update()
            return
        # If the data is the same as before(not changed)
        if inside_structure == options:
            notification('info', '設定は変更されていません。')
            return

        # Reflect the changes to the public variables
        #print(editor)
        TAG_COLOR = inside_structure['マーク色スキーム']
        wip_option = inside_structure['仕掛検索デフォルト']

        setting_cluster = {
            'tag': TAG_COLOR,
            'wip': wip_option
        }
        setting_str = json.dumps(setting_cluster, ensure_ascii=False)
        result = _DO.save_setting(USER_INS.uid, setting_str)
        if result[0] == 'e':
            notification('negative', f'設定の保存に失敗しました、処理を続行します。\nエラー詳細: {result[1]}\n設定は保存されていません。')
            USER_INS.retry_log(_DO.save_setting, USER_INS.uid, setting_str)
            return

        notification('positive', '設定変更を保存しました、新規タブから適用されます。')

    async def add_new_color():
        data = await editor.run_editor_method('get')
        tag_color: Dict[str, Dict[str, str]] = data['json']
        if '//新規（名前を変更してください' in tag_color['マーク色スキーム']:
            notification('negative', '1回では1つのマークしか追加できません、すでに追加されたマークを名前を希望するものに変更してから再度追加してください。')
            return
        tag_color['マーク色スキーム']['//新規（名前を変更してください'] = '#FFFFFF'
        editor.properties['content']['json'] = tag_color
        #print(editor)
        #editor.update()
    
    save.on_click(save_setting)
    newcolor.on_click(add_new_color)
    dialog.open()
    spinner_close()

async def excel_read_menu(menu_con: CustomExpansion, input_box: ui.textarea):
    from modules import ActiveBook, get_all_books_act
    ab = ActiveBook()

    async def on_book_change(e):
        if not e.value:
            return
        await ab.set_target_book(e.value)
        sheets = await ab.get_sheets()
        sheet_select.set_options(sheets)

    async def update_text_area(text):
        input_box.set_value('\n'.join(text))

    @spinner_deco
    async def read_books():
        if ab.book_name is None:
            notification('negative', '選択中のブックはありません。')
            return
        if not sheet_select.value:
            notification('negative', '選択中のシートはありません。')
            return
        data = await ab.read(sheet_select.value, columnname.value, duplicate=duplicatecheck.value)
        await update_text_area(data)

    if menu_con.value:
        try:
            with menu_con:
                with ui.column().classes('w-full gap-2 p-2'):

                    # ── Section header (matches _sec() in write dialog) ──────
                    with ui.row(align_items='center').classes('gap-2 mb-1'):
                        ui.element('div').style(
                            'width:3px; height:12px; background:#37a660; '
                            'border-radius:3px; flex-shrink:0'
                        )
                        ui.label('読込元').style(
                            'font-size:10px; font-weight:700; color:#37a660; '
                            'text-transform:uppercase; letter-spacing:0.07em'
                        )

                    # ── Book selection ───────────────────────────────────────
                    ui.label('Excel ブック').style(
                        'font-size:11px; font-weight:600; color:#6b7280'
                    )
                    book_names, active = [], None
                    try:
                        result = await get_all_books_act()
                        if result:
                            book_names, active = list(result[0]), result[1]
                    except Exception:
                        pass
                    book_select = (
                        ui.select(book_names, value=active)
                        .classes('w-full').props('outlined dense options-dense')
                    )
                    book_select.on_value_change(on_book_change)

                    # ── Sheet selection ──────────────────────────────────────
                    ui.label('シート').style(
                        'font-size:11px; font-weight:600; color:#6b7280'
                    )
                    initial_sheets = []
                    if active:
                        try:
                            await ab.set_target_book(active)
                            initial_sheets = await ab.get_sheets()
                        except Exception:
                            pass
                    sheet_select = (
                        ui.select(initial_sheets, label='')
                        .classes('w-full').props('outlined dense options-dense')
                    )

                    # ── Column name input ────────────────────────────────────
                    ui.label('列名').style(
                        'font-size:11px; font-weight:600; color:#6b7280'
                    )
                    columnname = ui.input(
                        placeholder='例: A',
                        validation={
                            'A–Z で入力してください':
                                lambda v: bool(re.match(r'^[a-zA-Z]{1,3}$', v))
                        },
                    ).classes('w-full').props('outlined dense')

                    # ── Options + Read button ────────────────────────────────
                    with ui.row(align_items='center').classes('w-full justify-between mt-1'):
                        duplicatecheck = ui.checkbox('重複排除').classes('text-xs')
                        ui.button('読込', icon='download', on_click=read_books).props(
                            'color=positive dense'
                        )

        except Exception as e:
            print(e)
    else:
        menu_con.clear()
            
def call_settings():
    #spinner_open()
    setting_screen()

def server_shutdown():
    def shutdown_message():
        description.content = 'この画面が自動的に閉じられていない場合は、手動でこの画面を閉じてください。'
        confirm.disable()
        cancel.disable()
        app.shutdown()
    with ui.dialog().classes('w-full').props('backdrop-filter="blur(8px) brightness(40%)"').props('persistent') as dialog:
        with ui.card(align_items='center'):
            with ui.row(align_items='start').classes('w-full bg-red h-12'):
                ui.label('プログラム停止').classes('text-white text-lg font-semibold')
            with ui.row():
                with ui.card():
                    description = ui.html('下記プロセス終了ボタンを押してアプリを終了。（あるいはこの画面を閉じればプロセスが終了になります)', sanitize= False)
            with ui.row(align_items='center'):
                confirm = ui.button('プロセス終了', on_click=shutdown_message)
                cancel = ui.button('取消', on_click= dialog.close)
    
    dialog.open()

def update_shutdown():
    def shutdown_message():
        description.content = 'この画面が自動的に閉じられていない場合は、手動でこの画面を閉じてください。'
        confirm.disable()
        cancel.disable()
        import subprocess
        subprocess.Popen([r"\\sflise02\伊勢原２\接合・セラミック部\業務課\2.業務G\Apps\プロセスマネージャー\install.exe"])
        app.shutdown()
    with ui.dialog().classes('w-full').props('backdrop-filter="blur(8px) brightness(40%)"').props('persistent') as dialog:
        with ui.card(align_items='center'):
            with ui.row(align_items='start').classes('w-full bg-green h-12'):
                ui.label('新バージョン').classes('text-white text-lg font-semibold')
            with ui.row():
                with ui.card():
                    description = ui.html(f'新バージョンがあります。アップデートしますか？<br>"はい"を押すと現在のプロセスを終了にし、アップデートプロセスに移します。', sanitize= False)
            with ui.row(align_items='center'):
                confirm = ui.button('アップデート', on_click=shutdown_message)
                cancel = ui.button('取消', on_click= dialog.close)
    
    dialog.open()

async def client_diagnose(client: Client, tag_str: str = 'client diagnose'):
    print(tag_str)
    print(f'log start {datetime.now()} ----------------\n\n')

    alert = f"""
    new Promise(resolve => {{
        resolve(document.visibilityState === 'visible');
            }})
        """
    
    allc = Client.instances

    for c_id, c_ins in allc.items():
        print(f'######################\n\nid = {c_id}\n\nins = {c_ins}1111111111111111')
        print(f'socket = {c_ins.has_socket_connection}\nwaiting for disconnection = {c_ins.is_waiting_for_disconnect}\nexistence = {c_ins.check_existence()}\nelement = \n\n{c_ins.content}')
        print(f'type of cins contents: {type(c_ins.content)}, bool of cins contents: {True if c_ins.content else False}, content : {c_ins.elements}, session label = {True if 'セッション管理をリセット' in str(c_ins.content) else False}')
        if c_ins.content:
            try:
                a = await c_ins.run_javascript(alert)

                print(f'WITH CONTENTS VISIBLE ===== {a}')
            
            except:
                pass

            continue

        if c_ins == client:
            print(f'found current client : {c_ins}')

            #a = await c_ins.run_javascript(alert)

            #print(f'CURRENT VISIBLE ===== {a}')

    print(f'log end ----------------\n\n')  

async def check_duplicated_page(client: Client, app_name: str) -> bool:
    document_vis_checker = """
    new Promise(resolve => {
        resolve(document.visibilityState === 'visible');
            })
        """
    helper_message = ''
    instrction = ''
    all_clients = Client.instances
    main_page_signal = False
    #print('\n Start of test log\n⇓⇓⇓⇓⇓⇓⇓⇓⇓⇓⇓⇓⇓⇓\n')
    try:
        for client_instance in list(all_clients.values()):
            for e in list(client_instance.elements.values()):
                if isinstance(e, ui.header):
                    header: ui.header = e
                    row: ui.row = e.default_slot.children[0]
                    for item in row.default_slot.children:
                        if isinstance(item, ui.label):
                            label: ui.label = item
                            
                            if label.text.split()[0] in app_name:
                                try:
                                    javascript_signal = await client_instance.run_javascript(document_vis_checker)
                                    main_page_signal = True
                                    if javascript_signal:
                                        helper_message = '実行中のインスタンスは別のブラウザから開かれている可能性があります。'
                                        instrction = '開いているほかのブラウザから’プロセスマネージャー’という名前のタブを探して、その**タブからアクセス**してください。\n**このブラウザを閉じてください。**'
                                    else:
                                        helper_message = '実行中のインスタンスは別のタブから開かれている可能性があります。'
                                        instrction = '開いているブラウザのタブから’プロセスマネージャー’という名前のタブを探して、その**タブからアクセス**してください。\n**このタブを閉じてください。**'
                                    #print(f'client_instance found with mp:\n  -> {client_instance}\n -> {main_page_signal}')
                                except:
                                    #print('cant get main page signal')
                                    pass

                                break
                            #else:
                                #print(f'no mp found')
                        #else:
                            #print('no lable instance exists')
                #else:
                    #print('no header exists')
    except RuntimeError as runtime_error:
        #check if the runtime error is caused by the change of the elements ditionary during the iteration,
        if 'dictionary keys changed during iteration' in str(runtime_error) or 'dictionary changed size during iteration' in str(runtime_error) :
            notification('warning', '再接続の発生など、元のセッションより再開できない問題が発生しましたので、リスタートします。')
            import asyncio
            for i in range(5):
                await asyncio.sleep(1)
                notification('info', f'リセット{5-i}秒前...')
            notification('info', 'リセット開始...')
            try:
                spinner_close()
            except Exception:
                pass
            global TAB_INITIALSED
            TAB_INITIALSED = False
            ui.navigate.reload()
    except Exception as e:
        #broad guard: catch any unexpected error during page duplication check and reload gracefully
        print(f'Error during duplicated page check: {e}')
        try:
            spinner_close()
        except Exception:
            pass
        ui.navigate.reload()
    
    if main_page_signal:
        ui.markdown(f'''
## 既存のブラウザまたはタブを探してください

{helper_message}

新しいウィンドウやタブで同時に複数起動することはできません。

**ご利用中の画面やタブを探して、そちらから操作を続けてください。**

{instrction}

- 既存のウィンドウやタブが見つからない場合は、すべてのブラウザを一度閉じてから再度アクセスしてください。
- それでも解決しない場合は、システム管理者にご連絡ください。        
''')
        #reset = ui.button('セッション管理をリセット', on_click=lambda session = app.storage.general[USER_INS.uid]: reset_session(session))

        ui.notify(type='warning', message=f'{helper_message}開いている画面からアクセスしてください。', close_button= True, position='center')
        #await client_diagnose(client, 'inside session holder')
        return True
        #return True
    
    else:
        return False
    
def about_screen():
    import random
    glow_class = f'glow-{random.randint(1000, 9999)}'

    with ui.dialog().props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:

        with ui.element('div').style(f'''
            animation: {glow_class} 2s ease-in-out infinite;
            background: #FFFFFF;
            color: black;
            padding: 24px;
            border-radius: 8px;
            box-shadow: 0 0 20px #9333ea;
        '''):

            ui.label(f'{APP_NAME} {APP_VERSION}').classes('text-center bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-2xl font-extrabold')
            
            ui.space().classes('h-8')
            
            ui.html('''
            <style>
            .tech-logo-container {
                display: inline-block;
                background: #000000;
                padding: 1.5rem 3rem;
                border-radius: 0.5rem;
                box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
            }
            .tech-logo-text {
                color: white;
                font-family: 'Courier New', monospace;
                font-size: 2.35rem;
                font-weight: bold;
                letter-spacing: 0.1em;
                margin: 0;
            }
            </style>
            <div class="tech-logo-container">
                <h1 class="tech-logo-text">Process Manager</h1>
            </div>
            ''',
            sanitize=False)

            ui.space().classes('h-8')

            ui.label('By Richard Chang').classes('text-center bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg')
            ui.space().classes('h-8')
            with ui.label().classes('text-center bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-lg text-grey'):
                ui.markdown("""
                ### License Notice

                This program uses the following open-source libraries:

                - **NiceGUI** (MIT License)  
                https://github.com/zauberzeug/nicegui  
                Copyright © Zauberzeug GmbH

                - **pandas** (BSD 3-Clause License)  
                https://github.com/pandas-dev/pandas  
                Copyright © The pandas development team

                - **pywin32** (MIT License)  
                https://github.com/mhammond/pywin32  
                Copyright © Mark Hammond

                - **wonderwords** (MIT License)  
                https://github.com/saffsd/wonderwords  
                Copyright © Sam Forde

                - **oracledb** (Apache License 2.0)  
                https://github.com/oracle/python-oracledb  
                Copyright © Oracle and/or its affiliates

                - **pyperclip** (BSD License)  
                https://github.com/asweigart/pyperclip  
                Copyright © Al Sweigart

                Other dependencies may be present; please refer to their respective documentation for license details.
                """)
            
            # Inject the keyframes for this specific element

            ui.html(f'''
            <style>
                @keyframes {glow_class} {{
                    0%, 100% {{ 
                        box-shadow: 
                            0 0 5px #9333ea,
                            0 0 10px #9333ea,
                            0 0 15px #9333ea,
                            0 0 20px #9333ea,
                            0 0 35px #9333ea;
                    }}
                    50% {{ 
                        box-shadow: 
                            0 0 20px #9333ea,
                            0 0 30px #9333ea,
                            0 0 40px #9333ea,
                            0 0 60px #9333ea,
                            0 0 80px #9333ea;
                    }}
                }}
            </style>
            ''',
            sanitize=False)

    dialog.open()

async def reset_fallback(error_message: str):
    notification('negative', 'セッションの状態が不安定です、セッション管理をリセットします。')
    notification('info', f'エラー内容: {error_message}', multiline=True)
    import asyncio
    for i in range(5):
        await asyncio.sleep(1)
        notification('info', f'リセット{5-i}秒前...')
    notification('info', 'リセット開始...')
    try:
        spinner_close()
    except Exception:
        pass
    ui.navigate.reload()
    


#MARK: client_instance Starter

async def _fetch_and_store_user_name(uid: str) -> str | None:
    """Query Oracle for the user's display name and persist it in SQLite.
    Returns the name string if found, None on any failure — never raises."""
    try:
        rows = await M.OC.asy_q_execute(
            'SELECT USER_NAME FROM MES_USER_INFORMATION_LC WHERE USER_CODE = :1',
            (uid,)
        )
        if rows:
            name = rows[0][0]
            if name:
                _DO.update_user_name(uid, str(name))
                return str(name)
    except Exception:
        pass
    return None


@ui.page('/')
async def create_app(client: Client):
    ui.add_body_html(_GANTT_CSS)
    ui.on('tabvisible', KEY_CACHE.reset)
    ui.add_head_html('''
                 <style>
                        .selected-column {
                         background-color: #00DEC4;
                         color: #003A33;
                         font-weight: bold;
                         }
                        .selected-column:hover {
                         background-color: #89FFF1 !important;
                         color: #003A33;
                         }
                        .selected-column-table {
                         color: #D6D6D9;
                         font-weight: bold;
                         }
                        .selected-column-table:hover {
                         background-color: #ECECED !important;
                         color: #D6D6D9;
                         }
                    </style>


                <script>
                    document.addEventListener('visibilitychange', () => {
                        if (document.visibilityState === 'visible') {
                            emitEvent('tabvisible');
                        }
                    });
                </script>
                ''')
    ui.add_css(
                '''
                .q-separator--vertical.nicegui-separator {
                    width: 1px;
                }
                .q-separator--horizontal.nicegui-separator {
                    width: 100%;
                }
                ''',
                shared=True
            )

    try:
        await client.connected()

        global USER_INS
        USER_INS = UserInstance()
        #print(f'TEST MODE: {TEST}')
        if not TEST:
            duplicate = await check_duplicated_page(client, APP_NAME)

            if duplicate:
                return

        spinner_create()
        if not _DO.check_user_exists(USER_INS.uid):
            default_setting = {'tag' : TAG_COLOR, 'wip' : wip_option}
            json_str = json.dumps(default_setting, ensure_ascii=False)
            result = _DO.new_user_single(USER_INS.uid, json_str)
            if result[0] == 'e':
                USER_INS.retry_log(_DO.new_user_single, USER_INS.uid, default_setting)
                notification('negative', f'新しいユーザーの初期化が失敗しました、そのままアプリを使用できますが、設定の変更などは保存されません。\nエラー詳細:{result[1]}', multiline= True)
            elif result[0] == 's':
                fetched_name = await _fetch_and_store_user_name(USER_INS.uid)
                display = f'{result[1]} {fetched_name}' if fetched_name else result[1]
                notification('positive', f'>ユーザー {display} 初期化しました。')
        else:
            USER_INS.initial_user()
            if not USER_INS.user_name or USER_INS.user_name == '-':
                async def _bg_fetch():
                    fetched = await _fetch_and_store_user_name(USER_INS.uid)
                    if fetched:
                        USER_INS.user_name = fetched
                asyncio.ensure_future(_bg_fetch())
                display_name = USER_INS.uid
            else:
                display_name = USER_INS.user_name
            notification('positive', f'👤ユーザー {display_name} ')
    except Exception as e:
        tb_str = ''.join(traceback.format_exception(type(e), e, e.__traceback__))
        _app_logger.error(f'create_app error: {tb_str}')
        print(tb_str)
        notification('negative', '初期化時にエラーが発生しました、再度アクセストライしても解消しない場合は、アプリ保存場所にある"再開できない場合"のバッチファイルを実行してください')
        await reset_fallback(str(e))
        return

    async def recreate_session() -> bool:
        if len(current_tabs) == 1:
            return False
        recreate_spinner.open()
        recreate_spinner.set_visibility(True)
        try:
            last_tab = None
            cached_tabs = dict(current_tabs)
            current_tabs.clear()
            for key in cached_tabs:
                if key != 'default_0':
                    datacluster = cached_tabs[key][0]
                    if datacluster is None:
                        return False
                    last_tab = await newtab(*datacluster)
            tabs.set_value(last_tab[0])
            search_prompt['prompt'] = False
            return True
        except Exception as e:
            tb_str = ''.join(traceback.format_exception(type(e), e, e.__traceback__))
            _app_logger.error(f'recreate_session error: {tb_str}')
            print(tb_str)
            notification('negative', 'セッション再開中にエラーが発生しました。')
            return False


    #cache = app.storage.client
    # Create the main layout

    with ui.dialog().props('persistent') as recreate_spinner, ui.card():
        ui.label('セッション再開中...')
        ui.spinner(size= 'lg')
    recreate_spinner.set_visibility(False)    

    
    #ui.header().classes('bg-grey-500 rounded-lg transition-transform duration-300 hover:scale-110 cursor-pointer ').style('height: 8px !important; min-height: 8px !important;')
    with ui.header().classes(f'flex justify-between items-center bg-[{ACCENT_COLOR}] text-white p-4'):
        #TEST = False
        #if not TEST:
        #    with ui.row().classes('ml-auto'):
        #        def close():
        #            app.native.main_window.destroy()
        #            app.shutdown()
        #        minimize = ui.button(icon='minimize', on_click=
        #                             lambda: app.native.main_window.minimize()).props('flat text-color="white" dense')
        #        fullscreen = ui.button(icon='fullscreen', on_click=
        #                               lambda: app.native.main_window.toggle_fullscreen()).props('flat text-color="white" dense')
        #        close = ui.button(icon='close', on_click=close).props('flat text-color="white" dense')
        with ui.row().classes('w-full flex justify-between items-center'):
            # Left side - Menu button
            ui.button('', icon='menu').props('flat').classes('text-white').on('click', lambda: left_drawer.toggle())
            
            # Center - App name and version
            ui.label(f'{APP_NAME} - v{APP_VERSION}').classes('text-xl font-bold')
            
            # Right side - Search button
            ui.button('検索', icon='search').props('flat').classes('text-white').on('click', open_search_popup)
    
    # Create left drawer for menu
    with ui.left_drawer(elevated=True).props('overlay=True').classes(f'bg-[{SUB_ACCENT_COLOR}]/50').props('bordered') as left_drawer:
        menu_title = ui.label('メニュー').classes('text-lg font-bold p-4')
        with ui.list().classes('w-full') as drawer_list:
            memo = ui.item('メモ帳', on_click= lambda t= menu_title, l= drawer_list, d= left_drawer: memo_drawer(t, l, d))
            relation = ui.item('品番/品目関係検索', on_click= lambda t= menu_title, l= drawer_list, d= left_drawer: sap_relation_quick_drawer(t, l, d))
            flow = ui.item('フロー/工程検索', on_click=flow_and_process_code)
            _dataviewer = ui.item('データビュアー', on_click=dataviewer)
            ui.separator().classes('w-full')
            if not TEST:
                _sched = ui.item('仕入先スケジュール', #icon='calendar_month',
                                on_click=lambda: supplier_sched_drawer(menu_title, drawer_list, left_drawer))
            ui.separator().classes('w-full')
            settings = ui.item('設定', on_click=call_settings)
            #tutorial_b = ui.item('使用方法', on_click=tutorial)
            ui.separator().classes('w-full')
            ui.item('About', on_click=about_screen)
            ui.separator().classes('w-full')
            ui.item('終了', on_click=server_shutdown)
            ui.separator().classes('w-full')
            ui.item('倉庫係システム(開発中)', on_click=lambda: ui.navigate.to('/tp'))
            #ui.item('test_transcation_screen', on_click=lambda: ui.navigate.to('/transcation'))
            
            
            
    
    left_drawer.hide()
    # Main content area with tabs
    global tabs
    tabs = ui.tabs().classes('w-full rounded-xl shadow-xl')#.bind_visibility(search_prompt, 'prompt', value=False)
    #await ui.context.client.connected()
    #tabs.bind_value(app.storage.tab, 'tabs')
    # Tab panels container

    search_prompt['prompt'] = True
    global tab_panels
    tab_panels = ui.tab_panels(tabs).classes('w-full h-full')
    tab_panels.on_value_change(lambda _: ui.run_javascript(
        "window._aggridResizeHandler && window._aggridResizeHandler()"))
    # Create default tab with prompt

    if len(current_tabs) > 1:
        #print(current_tabs)
        fin = await recreate_session()
        if fin:
            search_prompt['prompt'] = False
            recreate_spinner.close()
            recreate_spinner.set_visibility(False)
        else:
            await reset_fallback('セッションの再開に失敗しました。')


    # Pre-warm ag-grid JavaScript bundle so browser parses it before first query
    with ui.element('div').style('display:none'):
        ui.aggrid({'columnDefs': [], 'rowData': []})

    with tabs:
        default_tab = ui.tab('Welcome').classes('rounded-t-lg').props('icon=home').bind_visibility(search_prompt, 'prompt')
    
    with tab_panels:
        with ui.tab_panel(default_tab).style('''
                            position: relative;
                            padding: 20px;
                            height: 100%;
                            border-radius: 12px;
                            background-color: #e3f2fd;
                            font-family: sans-serif;
                            overflow: hidden;
            ''') as default_tab_panel:
            with ui.card().classes('w-full h-full flex justify-center items-center'):
                ui.label('検索ボタンをクリックして新規検索を実行').classes('text-xl text-gray-500')
                pass  # drawer entry point: 仕入先スケジュール
            #with ui.card().classes('w-full h-full flex justify-center items-center'):
                #ui.label('クイックアクセス').classes('text-sm')
                #with ui.row().classes('w-full'):
                #    add_quick = ui.button('検索パターン追加')
                #    ui.separator().props('vertical')
                #    empty_place_holder = ui.label('追加ボタンをクリックしてよく使う検索パターンを追加').classes('text-sm text-gray-500')
            with ui.card().classes('w-full h-full flex justify-center items-center'):
                ui.html(f'<strong>Beta Test Build.</strong><br>ユーザー {USER_INS.uid}', sanitize=False)
            ui.icon('troubleshoot').classes('absolute').style('''
                right: -20px;
                bottom: -40px;
                font-size: 215px;
                color: rgba(1, 1, 1, 0.2);
                pointer-events: none;
            ''')

            #with ui.card() as card:
            #    test = ui.button('test')
            #    gantt_container = ui.column()

            #    async def testfunc():
            #        try:
            #            from test_mod import fetch_from_df
            #            lots = await fetch_from_df()

            #            from customwidget import Gantt
            #            gantt = Gantt()
            #            gantt.create_gantt_from_raw_data(
            #                lots, 
            #                container=gantt_container,
            #                page_size=80,  # rows per page
            #                cell_width=80
            #            )
            #        except Exception as e:
            #            import traceback
            #            traceback.print_exc()
                
            #    test.on_click(testfunc)
                
    

    '''else:
        print(current_tabs)
        

        for k in current_tabs:
            if k != 'default_0':
                current_tabs[k][0].move(tabs)
        for k in current_tabs:
            if k != 'default_0':
                current_tabs[k][1].move(tab_panels)'''

    current_tabs['default_0'] = (default_tab, default_tab_panel)
    if len(current_tabs) == 1:
        tabs.set_value(default_tab)

    global keyb_listener
    keyb_listener = ui.keyboard(on_key= key_listener)

    #ui.add_head_html('''
    #                <style>
    #                    .selected-column {
    #                     background-color: #00DEC4;
    #                     color: #003A33;
    #                     font-weight: bold;
    #                     }
    #                    .selected-column:hover {
    #                     background-color: #89FFF1 !important;
    #                     color: #003A33;
    #                     }
    #                    .selected-column-table {
    #                     color: #D6D6D9;
    #                     font-weight: bold;
    #                     }
    #                    .selected-column-table:hover {
    #                     background-color: #ECECED !important;
    #                     color: #D6D6D9;
    #                     }
    #                </style>
    #                     ''')

    
    #await client_diagnose(client, 'after initailsed')

    #await client.disconnected()
    #print(f'\nuser storage = {app.storage.user}\ngeneral = {app.storage.general}')
    #app.storage.general[USER_INS.uid].clear()
    if not version_check():
        update_shutdown()



def tutorial():
    with ui.dialog().classes('w-full h-5/6').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:
        with ui.column().classes('w-full h-3/4'):
            with ui.row().classes('w-full h-full gap-4'):
                with ui.card().classes('flex-1'):
                    ui.item('1')
        with ui.column().classes('w-full h-3/4'):
            with ui.card().classes('w-full h-full'):
                ui.label('2')
    dialog.open()
    

#MARK: Search screen function
def open_search_popup():
    # ISSUE3: Related to ISSUE1 in module file, can we implement a way to check and start connection to database when
    # user open the search popup
    import asyncio

    def modechange():
        search_mode['lot'] = True if search_mode['value'] in 'it' else False
        print(search_mode)

    def test_check():
        print(wip_option)

    with ui.dialog().classes('w-[28rem]').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog, ui.card().classes('w-96'):
        ui.label('検索').classes(f'text-lg font-bold mb-2 w-full')
        
        # Mode selection - create without initial value and set it after creation
        
        mode = ui.toggle(
            options= 
            {'i' : '指図進捗', 
             'w' : '仕掛状況', 
             't' : '関連指図'},
             value= 'i'
             ).classes('w-full justify-center mb-4').bind_value(search_mode).on_value_change(modechange)
                   
        textinput = ui.textarea(placeholder='検索').props('clearable').classes('w-full shadow-xl rounded-lg focus: border-4 border-indigo-200 border-b-indigo-500')
        
        

        # Container for detail options (conditionally visible)
        detail_container = ui.element('div').classes('w-full')
        
        # Process mode details
        with detail_container:
            ui.label('オプション').classes('font-bold mt-2 shadow-lg')
            ui.separator()
            with ui.element('div').bind_visibility_from(search_mode, 'value', value='w'):
                ui.label('工場:').classes('pt-2')
                ui.checkbox('伊勢原工場').bind_value(wip_option, 'ise').on_value_change(test_check)
                ui.checkbox('宮田工場').bind_value(wip_option, 'miya').on_value_change(test_check)
                ui.separator()
                ui.checkbox('工程開始前の指図を表示').bind_value(wip_option, 'prestart').on_value_change(test_check)

            with ui.element('div').bind_visibility_from(search_mode, 'lot', value=True):
                read_from_excel = CustomExpansion('Excelブックから読込', on_value_change=lambda:excel_read_menu(read_from_excel, textinput))
                
                    
            
            with ui.element('div') as like_container:
                ui.separator()
                with ui.scroll_area() as like_scroll:
                    like_skeleton = ui.skeleton('text').classes('w-full text-subtitle1')
                    
            like_skeleton.set_visibility(False)
            like_container.set_visibility(False)

        def clear_search_cache():
            like_scroll.clear()
            LIKE_SEARCH.clear()

        _qos_token = [0]  # cancellation counter; incremented to invalidate in-flight queries

        async def replace_complete(key: str, replacement: str, event = None):
            searchkey = key.replace('%', '*')
            inputed: str = textinput.value
            if not inputed:
                return
            replaced = inputed.upper().replace(searchkey, replacement)
            textinput.set_value(replaced)
            textinput.update()
            del LIKE_SEARCH[key]
            like_scroll.clear()
            like_container.set_visibility(False)
            await toggle_change()

        async def query_on_search_():
            global LIKE_SEARCH
            global ON_TYPE_SEARCH_ALERT
            if textinput.value is not None and "*" in textinput.value and mode.value == 'w':
                if not like_container.visible or ON_TYPE_SEARCH_ALERT:
                    # Show skeleton inside container before awaiting so user sees a loading indicator
                    # (skeleton lives inside like_container, so container must be visible too)
                    like_skeleton.set_visibility(True)
                    like_container.set_visibility(True)
                    if ON_TYPE_SEARCH_ALERT:
                        clear_search_cache()
                    _qos_token[0] += 1          # invalidate any previous in-flight query
                    my_token = _qos_token[0]    # capture this query's identity
                    try:
                        result = await query_on_search(textinput, None, LIKE_SEARCH.keys())
                    except E.NoKeyExc:
                        if _qos_token[0] == my_token:
                            like_skeleton.set_visibility(False)
                            like_container.set_visibility(False)
                        notification('negative', '有効なワイルドカードキーはありません：ワイルドカード検索するには、最低限"DB" + 数字3桁 + "*"で入力してください。', multiline=True)
                        return
                    except Exception as e:
                        if _qos_token[0] == my_token:
                            like_skeleton.set_visibility(False)
                            like_container.set_visibility(False)
                        warningtext = WARNING_NOTIFICATION['UnExp']
                        notification(warningtext[0], f'{warningtext[1]}: {e}')
                        return
                    # Discard stale result if user already cancelled
                    if _qos_token[0] != my_token:
                        like_skeleton.set_visibility(False)
                        like_container.set_visibility(False)
                        return
                    LIKE_SEARCH |= result
                    top = tuple(LIKE_SEARCH.keys())[0] if LIKE_SEARCH else 'EE'
                    with like_scroll:
                        if top == 'EE':
                            like_skeleton.set_visibility(False)
                            like_container.set_visibility(False)
                            return
                        like_r = LIKE_SEARCH[top]
                        if not isinstance(like_r, M.query_):
                            top_label = ui.label(f'{top[:-1]}_ワイルドカード').classes('bg-[#5994A7] w-full rounded-lg font-semibold pl-2')
                            for key in like_r:
                                if key == 'DBのワイルドカード検索はできません、':
                                    ON_TYPE_SEARCH_ALERT = True
                                namelabel = ui.label(key).classes('bg-[#FFFFFF] hover:bg-[#ECECEC]  w-full rounded-lg pl-2').on('click', lambda k=top, r=key[0]: replace_complete(key=k, replacement=r))
                            like_skeleton.set_visibility(False)  # hide after loop — covers empty result too
                        else:
                            # result is still a query_ placeholder (e.g. FrequentExc with no cache)
                            like_skeleton.set_visibility(False)
                            like_container.set_visibility(False)

            elif not textinput.value or '*' not in textinput.value:
                _qos_token[0] += 1             # cancel any in-flight query
                like_skeleton.set_visibility(False)  # hide immediately — don't wait for result
                like_container.set_visibility(False)
                clear_search_cache()

        async def toggle_change():
            if mode.value == 'w':
                await query_on_search_()
            else:
                clear_search_cache()
                like_container.set_visibility(False)

        # Action buttons
        with ui.row().classes('w-full justify-end gap-2 mt-4'):
            ui.button('取消', icon='close').props('flat').on('click', lambda: dialog.close())
            ui.button('検索', icon='search').props('color=primary').on('click', lambda: perform_search(textinput, dialog))

        textinput.on_value_change(query_on_search_)
        mode.on_value_change(toggle_change)

        asyncio.ensure_future(M.OC.ensure_connection())  # ISSUE3: warm up DB connection proactively
        dialog.open()

#MARK: AGGrid/New Tab related
#def get_selected_cols(rows: List[dict], id):
#    if any(tuple(v.highlight fo#r _, v in AGGRID_SELECTION[id].i#tems() if AGGRID_SELECTION[id])):
#        return tuple({k : v for k, v in row.items() if k != 'Sele#cted' and AGGRID_SELECTION[id][k].vi#sible and AGGRID_SELECTION[id][k].highlight} for row in rows)
#    else:
#        return tuple({k : v for k, v in row.items() if k != 'Sele#cted' and AGGRID_SELECTION[id][k].visible} for row in rows)

    

#def deselect_all(id):
#   # for k in AGGRID_SELECTION[id]:#
#        AGGRID_SELECTION[id][k].highlight_flip(False)

def clear_event_cache():
    global CLICK_EVENT_CACHE 
    CLICK_EVENT_CACHE = None


def tab_delete(tab_obj: ui.tab, tab_id: int):
    #tabs.remove(tab_obj)
    #print(current_tabs, f'tab id == \n\n {tab_id}\n################')
    current_tabs[tab_id][2].delete()
    tab_obj.delete()
    current_tabs.pop(tab_id, "0")
    #print(len(current_tabs))
    if len(current_tabs) == 1:
        search_prompt['prompt'] = True

    last_tab_key_t = tuple(k for k in current_tabs.keys() if k != 'default_0')
    if not last_tab_key_t:
        last_tab_key = 'default_0'
    else:
        last_tab_key = last_tab_key_t[-1]    
    #last_tab_key = tuple(current_tabs.keys())[-1]
    last_tab = current_tabs[last_tab_key][1]
    #print(f'\nlast tab key\n{last_tab_key}\n')
    tabs.set_value(last_tab)

def aggrid_update(grid: ui.aggrid, grid_options: dict):
    ui.run_javascript(f"""
    const grid = getElement('{grid.id}');
    if (grid && grid.api) {{
        grid.api.setGridOption('columnDefs', {json.dumps(grid_options['columnDefs'])});
    }}
    """)    
    #grid.run_method('setGridOption', 'columnDefs', grid_options['columnDefs'])

def aggrid_data_update(grid: ui.aggrid, data: dict):
    grid.run_grid_method('setGridOption', 'rowData', data)

def tick_all_tree(treelist: List[ui.tree], tick: bool):
    if tick:
        for t in treelist:
            t.tick()
    else:
        for t in treelist:
            t.untick()




    
    

#MARK: New tab function
# ─── Supplier Schedule Sample ────────────────────────────────────────────────

_SUPPLIER_SCHED_GROUPS = [
    {'name': 'ブラスト',  'prefix': 'blast'},
    {'name': '精密洗浄', 'prefix': 'clean'},
]

_MOCK_SUPPLIER_DATA = [
    {'lot_no': '4212340001', 'db_drawing': 'NHK-A001', 'cust_pn': 'CPN-001', 'product_name': 'シャフト A型',       'quantity': 50,  'blast_supply': '2026-03-10', 'blast_deadline': '2026-03-11', 'blast_reply': '2026-03-12', 'clean_supply': '2026-03-13', 'clean_deadline': '2026-03-14', 'clean_reply': '2026-03-15', 'note1': '',       'note2': '',       'note3': ''},
    {'lot_no': '4212340002', 'db_drawing': 'NHK-B002', 'cust_pn': 'CPN-002', 'product_name': 'カバー B型',         'quantity': 30,  'blast_supply': '2026-03-11', 'blast_deadline': '2026-03-12', 'blast_reply': '2026-03-13', 'clean_supply': '2026-03-14', 'clean_deadline': '2026-03-15', 'clean_reply': '2026-03-16', 'note1': '要確認',   'note2': '',       'note3': ''},
    {'lot_no': '4212340003', 'db_drawing': 'NHK-C003', 'cust_pn': 'CPN-003', 'product_name': 'ブラケット C型',     'quantity': 100, 'blast_supply': '2026-03-12', 'blast_deadline': '2026-03-13', 'blast_reply': '2026-03-14', 'clean_supply': '2026-03-15', 'clean_deadline': '2026-03-16', 'clean_reply': '2026-03-17', 'note1': '',       'note2': '急ぎ',    'note3': ''},
    {'lot_no': '4212340004', 'db_drawing': 'NHK-D004', 'cust_pn': 'CPN-004', 'product_name': 'リング D型',         'quantity': 200, 'blast_supply': '2026-03-13', 'blast_deadline': '2026-03-14', 'blast_reply': '2026-03-15', 'clean_supply': '2026-03-16', 'clean_deadline': '2026-03-17', 'clean_reply': '2026-03-18', 'note1': '',       'note2': '',       'note3': ''},
    {'lot_no': '4212340005', 'db_drawing': 'NHK-E005', 'cust_pn': 'CPN-005', 'product_name': 'ピン E型',           'quantity': 75,  'blast_supply': '2026-03-14', 'blast_deadline': '2026-03-15', 'blast_reply': '2026-03-16', 'clean_supply': '2026-03-17', 'clean_deadline': '2026-03-18', 'clean_reply': '2026-03-19', 'note1': '',       'note2': '',       'note3': '在庫調整'},
    {'lot_no': '4212340006', 'db_drawing': 'NHK-F006', 'cust_pn': 'CPN-006', 'product_name': 'ハウジング F型',     'quantity': 20,  'blast_supply': '2026-03-15', 'blast_deadline': '2026-03-16', 'blast_reply': '',            'clean_supply': '2026-03-18', 'clean_deadline': '2026-03-19', 'clean_reply': '',            'note1': '回答待ち', 'note2': '',       'note3': ''},
    {'lot_no': '4212340007', 'db_drawing': 'NHK-G007', 'cust_pn': 'CPN-007', 'product_name': 'スリーブ G型',       'quantity': 60,  'blast_supply': '2026-03-16', 'blast_deadline': '2026-03-17', 'blast_reply': '2026-03-17', 'clean_supply': '2026-03-19', 'clean_deadline': '2026-03-20', 'clean_reply': '2026-03-20', 'note1': '',       'note2': '',       'note3': ''},
    {'lot_no': '4212340008', 'db_drawing': 'NHK-H008', 'cust_pn': 'CPN-008', 'product_name': 'カップリング H型',   'quantity': 15,  'blast_supply': '2026-03-17', 'blast_deadline': '2026-03-18', 'blast_reply': '2026-03-19', 'clean_supply': '2026-03-20', 'clean_deadline': '2026-03-21', 'clean_reply': '2026-03-22', 'note1': '',       'note2': '要注意',  'note3': ''},
]


_sched_state_map:   dict[str, dict] = {}  # vendor_code → per-vendor state dict
_sched_vendor_tabs: dict[str, int]  = {}  # vendor_code → new_tab.id (dedup index)


def _sched_log(msg: str) -> None:
    """Append a timestamped line to supplier_sched_debug.log in the project root."""
    import os, datetime
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'supplier_sched_debug.log')
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    try:
        with open(log_path, 'a', encoding='utf-8') as _f:
            _f.write(f'[{ts}] {msg}\n')
    except OSError:
        pass


def _make_proc_field(proc_name: str) -> str:
    """Safe JS field-name prefix: replace any non-word char (slash, space, etc.) with underscore."""
    return re.sub(r'[^\w]', '_', proc_name)


_DATE_PICKER_EDITOR = '''class {
    init(p) {
        this.el = document.createElement('input');
        this.el.type = 'date';
        this.el.value = p.value || '';
        this.el.style.cssText = 'width:100%;height:100%;border:none;outline:none;padding:0 4px;font-size:inherit;cursor:pointer;box-sizing:border-box';
    }
    getGui() { return this.el; }
    afterGuiAttached() { this.el.focus(); try { this.el.showPicker(); } catch(e) {} }
    getValue() { return this.el.value; }
    isPopup() { return false; }
}'''


_CUST_PN_EDITOR = '''class {
    init(p) {
        const s = ((window._custPnMap||{})[p.data.db_drawing]||[]);
        this.el = document.createElement('input');
        this.el.type = 'text';
        this.el.value = p.value || (s[0]||'');
        this.el.style.cssText='width:100%;height:100%;border:none;outline:none;padding:0 4px;font-size:inherit';
        if (s.length) {
            const id='_cpndl_'+Date.now();
            const dl=document.createElement('datalist'); dl.id=id;
            s.forEach(v=>{ const o=document.createElement('option'); o.value=v; dl.appendChild(o); });
            document.body.appendChild(dl);
            this.el.setAttribute('list',id);
            this._dl=dl;
        }
    }
    getGui(){return this.el;}
    afterGuiAttached(){this.el.focus();this.el.select();}
    getValue(){return this.el.value;}
    destroy(){if(this._dl)this._dl.remove();}
    isPopup(){return false;}
}'''


def _generate_vendor_sched_col_defs(process_names: list) -> list:
    # JS editable predicates
    _pin_only      = '(params) => params.node.rowPinned === "top"'
    _pin_no_lot    = '(params) => params.node.rowPinned === "top" && !window._lotIdEditing'
    _exist_only    = '(params) => params.node.rowPinned !== "top"'
    _all_no_lot    = '(params) => params.node.rowPinned !== "top" || !window._lotIdEditing'
    col_defs = [
        {'field': 'lot_id',       'headerName': '指図番号', 'width': 130, 'pinned': 'left',
         ':editable': _pin_only,
         'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
        {'field': 'db_drawing',   'headerName': 'DB図番',   'width': 130,
         ':editable': _pin_no_lot,
         'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
        {'field': 'product_name', 'headerName': '品名',     'width': 160,
         ':editable': _pin_no_lot,
         'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
        {'field': 'cust_pn',      'headerName': '客先品番', 'width': 120,
         ':editable': _all_no_lot,
         'filter': 'agTextColumnFilter', ':cellEditor': _CUST_PN_EDITOR},
        {'field': 'quantity',     'headerName': '数量',     'width': 70,  'editable': True,
         'filter': 'agNumberColumnFilter'},
    ]
    _date_col = {':cellEditor': _DATE_PICKER_EDITOR, 'filter': 'agTextColumnFilter'}
    for proc in process_names:
        safe = _make_proc_field(proc)
        col_defs.append({
            'headerName': proc,
            'children': [
                {'field': f'{safe}_ship_date',          'headerName': 'NHK支給日', 'width': 110,
                 'editable': True, **_date_col},
                {'field': f'{safe}_nbd',                'headerName': '納期',      'width': 110,
                 'editable': True, **_date_col},
                {'field': f'{safe}_vendor_commit_date', 'headerName': '回答納期',  'width': 110,
                 'editable': True, **_date_col},
                {'field': f'{safe}_schedule_id',        'headerName': f'{proc}_ID', 'hide': True},
            ]
        })
    col_defs += [
        {'field': 'comment1',     'headerName': '備考1',   'width': 120, 'editable': True},
        {'field': 'comment2',     'headerName': '備考2',   'width': 120, 'editable': True},
        {'field': 'comment3',     'headerName': '備考3',   'width': 120, 'editable': True},
        {'field': 'lot_status',   'headerName': '工程状態', 'width': 150,
         'filter': 'agTextColumnFilter'},
        {'field': '_comment_sid',  'hide': True},   # tracks which schedule_id owns this row's comment
        {'field': '_row_key',      'hide': True},   # unique row identifier for getRowId
        {'field': '_orig_lot_id',  'hide': True},   # original Oracle lot_id (before '-' substitution)
    ]
    return col_defs


def _safe_str(val) -> str:
    """Return str, treating None/NaN as empty string."""
    if val is None:
        return ''
    try:
        if pd.isna(val):
            return ''
    except (TypeError, ValueError):
        pass
    return str(val)


def _safe_date(val) -> str:
    """Convert Oracle date / pandas Timestamp / NaN to 'YYYY-MM-DD' string."""
    if val is None:
        return ''
    try:
        if pd.isna(val):
            return ''
    except (TypeError, ValueError):
        pass
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m-%d')
    s = str(val)
    return s[:10] if s else ''


_REAL_LOT_RE = re.compile(r'^(?:42\d{8}|[KNS](?:\d|-){6,10})$')


def _pivot_schedule(df: pd.DataFrame) -> list:
    """Pivot long-format vendor schedule rows to grid rows.

    Real lots (matching _REAL_LOT_RE) are grouped by lot_id — multiple Oracle
    rows (one per process) are merged into a single grid row with per-process
    column groups.

    Non-real lots (placeholder lot_ids like '-', '1', empty) are each kept as
    a separate grid row, using the Oracle schedule_id as the unique group key.
    The displayed lot_id is '-' for all such rows; _orig_lot_id stores the
    original Oracle value for SQLite fallback lookups.
    """
    groups: dict = {}   # group_key -> row dict
    order:  list = []   # preserves insertion order

    _sched_log(f'[PIVOT] start: {len(df)} Oracle rows')
    for _, r in df.iterrows():
        orig_lot = str(r.get('lot_id') or '')
        proc     = str(r.get('process_name') or '')
        safe     = _make_proc_field(proc)
        sid_str  = str(r.get('schedule_id') or '')

        is_real = bool(_REAL_LOT_RE.match(orig_lot))
        if is_real:
            group_key     = orig_lot
            effective_lot = orig_lot
        else:
            group_key     = sid_str if sid_str else f'_nk_{len(order)}'
            effective_lot = '-'   # display '-' for all non-real / placeholder lots

        _sched_log(f'  [PIVOT] lot={orig_lot!r} proc={proc!r} sid={sid_str!r} real={is_real} key={group_key!r}')

        if group_key not in groups:
            row = {
                'lot_id':       effective_lot,
                '_orig_lot_id': orig_lot,
                '_row_key':     group_key,
                'db_drawing':   _safe_str(r.get('db_drawing')),
                'product_name': _safe_str(r.get('product_name')),
                'cust_pn':      '',
                'quantity':     1,
                'comment1': '', 'comment2': '', 'comment3': '',
                'lot_status':   _safe_str(r.get('lot_status')),
                '_comment_sid': '',
            }
            groups[group_key] = row
            order.append(group_key)
        else:
            row = groups[group_key]
            if not row['db_drawing']:
                row['db_drawing']   = _safe_str(r.get('db_drawing'))
            if not row['product_name']:
                row['product_name'] = _safe_str(r.get('product_name'))

        row[f'{safe}_ship_date']          = _safe_date(r.get('ship_date'))
        row[f'{safe}_nbd']                = _safe_date(r.get('nbd'))
        row[f'{safe}_vendor_commit_date'] = _safe_date(r.get('vendor_commit_date'))
        row[f'{safe}_schedule_id']        = sid_str

    result = [groups[k] for k in order]
    _sched_log(f'[PIVOT] done: {len(result)} grid rows from {len(df)} Oracle rows')
    return result


def _normalize_date(s: str):
    """
    Convert common date inputs → 'YYYY-MM-DD', or None if unrecognisable.
    Accepted formats: YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD, YYYYMMDD,
                      MM/DD/YYYY, MM-DD-YYYY.
    """
    s = (s or '').strip()
    if not s:
        return None
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d',
                '%m/%d/%Y', '%m-%d-%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None


async def _write_schedule_row(
    lot_id: str,
    vendor_code: str,
    process_code: str,
    ship_date,        # YYYY-MM-DD — Oracle TO_DATE('YYYY-MM-DD') mask
    nbd=None,         # YYYY-MM-DD or None
) -> bool:
    """Oracle MERGE for one lot+process. Returns True on success.
    Does NOT write to SQLite — caller is responsible for comment upsert if needed."""
    try:
        await M.fast_query(
            'vendor_schedule_insert',
            {'lot_id':       lot_id,
             'vendor_code':  vendor_code,
             'process_code': process_code,
             'ship_date':    ship_date,
             'nbd':          nbd},
            False, commit=True,
        )
        return True
    except Exception as exc:
        logger.error(f'[_write_schedule_row] {exc}')
        return False


async def _quick_add_to_schedule_dialog(
    lot_ids: list,
    vendor_code=None,
    on_save=None,
    vendor_name=None,
) -> None:
    """Quick-add one or more lots to the supplier schedule.
    vendor_code=None  → show supplier selector (standalone mode).
    vendor_code=<str> → supplier fixed, skip selector (split-screen mode).
    on_save           → async callable called after successful writes (for grid refresh).
    NOTE: SQLite comment fields (db_drawing, product_name, quantity) are NOT populated
          by this path — they will be blank in the schedule grid until the row is edited.
    """
    _vc        = [vendor_code]
    _process   = [None]
    _ship_date = [None]          # YYYY-MM-DD after _normalize_date

    # ── section header helper (matches existing codebase style) ──────
    def _sec(title: str):
        with ui.row(align_items='center').classes('gap-2'):
            ui.element('div').style(
                'width:3px;height:14px;background:#37a660;'
                'border-radius:3px;flex-shrink:0;'
            )
            ui.label(title).style(
                'font-size:11px;font-weight:700;color:#37a660;'
                'text-transform:uppercase;letter-spacing:0.07em;'
            )

    with ui.dialog().props(
        'backdrop-filter="blur(8px) brightness(40%)"'
    ) as dlg:
        with ui.card().style(
            'border-radius:16px;min-width:420px;padding:0;'
            'overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,.18);'
        ):
            # ── header ────────────────────────────────────────────────
            with ui.element('div').style(
                'background:linear-gradient(135deg,#569bbe,#3d7fa3);'
                'padding:12px 16px;display:flex;align-items:center;gap:10px;width:100%;'
            ):
                with ui.element('div').style(
                    'width:32px;height:32px;border-radius:50%;'
                    'background:rgba(255,255,255,.2);'
                    'display:flex;align-items:center;justify-content:center;flex-shrink:0;'
                ):
                    ui.icon('event_available').style('color:white;font-size:18px;')
                with ui.element('div'):
                    ui.label('支給日程表に追加').style(
                        'color:white;font-weight:700;font-size:14px;line-height:1.3;'
                    )
                    subtitle = f'{vendor_name}  ·  {len(lot_ids)}件選択中' if vendor_name else f'{len(lot_ids)}件選択中'
                    ui.label(subtitle).style(
                        'color:rgba(255,255,255,.6);font-size:11px;'
                    )

            # ── body ──────────────────────────────────────────────────
            with ui.element('div').style(
                'padding:20px;display:flex;flex-direction:column;gap:16px;'
                'width:100%;box-sizing:border-box;'
            ):

                # Step 1: vendor selector (standalone only) ───────────
                vendor_section = ui.element('div').classes('w-full')
                vendor_select = None
                if vendor_code is None:
                    with vendor_section:
                        _sec('仕入先')
                        vendor_select = ui.select(
                            {}, label='仕入先を選択'
                        ).classes('w-full').props('outlined dense options-dense')
                else:
                    vendor_section.set_visibility(False)

                # Lot ID list ─────────────────────────────────────────
                _sec('選択指図')
                with ui.element('div').style(
                    'max-height:80px;overflow-y:auto;display:flex;'
                    'flex-wrap:wrap;gap:4px;'
                ):
                    for lid in lot_ids:
                        ui.html(
                            f'<span style="background:#f0f4f8;border-radius:20px;'
                            f'font-size:11px;font-family:monospace;padding:2px 10px;">'
                            f'{lid}</span>',
                            sanitize=False,
                        )

                # Process + date ──────────────────────────────────────
                _sec('工程 / 日付')
                process_select = ui.select(
                    {}, label='工程を選択'
                ).classes('w-full').props('outlined dense options-dense')
                date_btn   = ui.button('日付を選択').props('outline color=primary')
                date_label = ui.label('').style('font-size:11px;color:#6b7280;')

                # Footer ──────────────────────────────────────────────
                footer_row  = ui.row().classes('justify-end gap-2 mt-2 w-full')
                spinner_row = ui.row().classes('justify-center w-full')
                with footer_row:
                    ui.button('キャンセル', on_click=dlg.close).props('flat')
                    add_btn = ui.button('追加').props('color=primary disable')
                with spinner_row:
                    ui.spinner(size='md')
                spinner_row.set_visibility(False)

    # ── date picker (separate dialog) ────────────────────────────────
    with ui.dialog() as date_dlg, ui.card().classes('p-2'):
        date_picker = ui.date().props('mask="YYYY/MM/DD"')
        def _confirm_date(_=None):
            if date_picker.value:
                _ship_date[0] = _normalize_date(date_picker.value)
                date_btn.set_text(date_picker.value)
            date_dlg.close()
            _refresh_add_btn()
        ui.button('確定', on_click=_confirm_date).props('color=primary flat')

    date_btn.on('click', date_dlg.open)

    # ── enable / disable add button ───────────────────────────────────
    def _refresh_add_btn():
        ok = bool(_vc[0] and _process[0] and _ship_date[0])
        if ok:
            add_btn.props(remove='disable')
        else:
            add_btn.props('disable')

    # ── load vendor list (standalone only) ───────────────────────────
    async def _load_vendors():
        df = await M.fast_query('NOKEY_vendor_master', [], True)
        if df is None or df.empty:
            ui.notify('仕入先データの取得に失敗しました', type='negative')
            dlg.close()
            return
        df.columns = df.columns.str.lower()
        options = {str(r['vendor_code']): str(r['vendor_name'])
                   for _, r in df.iterrows()}
        vendor_select.options = options
        vendor_select.update()

    # ── load process list for selected vendor ─────────────────────────
    async def _load_processes():
        if not _vc[0]:
            return
        df = await M.fast_query('vendor_processes', {'vendor_code': _vc[0]}, True)
        if df is None or df.empty:
            ui.notify('工程データの取得に失敗しました', type='negative')
            return
        df.columns = df.columns.str.lower()
        options = {str(r['process_code']): str(r['process_name'])
                   for _, r in df.iterrows()}
        process_select.options = options
        process_select.update()

    async def _on_vendor_change(e):
        _vc[0] = e.value
        await _load_processes()
        _refresh_add_btn()

    def _on_process_change(e):
        _process[0] = e.value
        _refresh_add_btn()

    process_select.on_value_change(_on_process_change)
    if vendor_select is not None:
        vendor_select.on_value_change(_on_vendor_change)

    # ── confirm: write all lots ───────────────────────────────────────
    async def _on_confirm():
        footer_row.set_visibility(False)
        spinner_row.set_visibility(True)
        ok_count   = 0
        fail_count = 0

        for lid in lot_ids:
            success = await _write_schedule_row(
                lid, _vc[0], _process[0],
                _ship_date[0], None,
            )
            if success:
                ok_count += 1
            else:
                fail_count += 1

        dlg.close()

        if fail_count == 0:
            ui.notify(f'{ok_count}件追加しました', type='positive')
        elif ok_count == 0:
            ui.notify('追加に失敗しました', type='negative')
        else:
            ui.notify(f'{ok_count}件追加 / {fail_count}件失敗', type='warning')

        if on_save is not None and ok_count > 0:
            await on_save()

    add_btn.on('click', _on_confirm)

    # ── open and load ─────────────────────────────────────────────────
    dlg.open()
    if vendor_code is None:
        await _load_vendors()
    else:
        await _load_processes()


async def _open_supplier_schedule_sample(params: dict):
    import asyncio as _asyncio
    vendor_code = params['vendor_code']
    vendor_name = params.pop('vendor_name', vendor_code)  # pop — not a SQL bind variable
    _sched_state = _sched_state_map.setdefault(vendor_code, {
        'grid': None, 'vendor_code': None, 'vendor_name': None,
        'toggle_form': None, 'params': None,
    })

    # Fetch process list and schedule data (full date range — no date filter applied)
    proc_df  = await M.fast_query('vendor_processes', {'vendor_code': vendor_code}, True)
    sched_df = await M.fast_query('vendor_schedule',  params, True)

    if sched_df is None:
        ui.notification('クエリエラーが発生しました', type='negative')
        return
    sched_df.columns = sched_df.columns.str.lower()

    # Deduplicate: the SQL join produces duplicate rows (same lot+process+sid).
    if not sched_df.empty and 'schedule_id' in sched_df.columns:
        sched_df = sched_df.drop_duplicates(
            subset=['lot_id', 'process_name', 'schedule_id'])

    # process_names from vendor_processes (PROMAN_VENDOR_PROCESS_TYPE order by PROCESS_CODE)
    # This defines column groups regardless of which lots appear in the date range.
    if proc_df is not None and not proc_df.empty:
        proc_df.columns = proc_df.columns.str.lower()
        process_names = proc_df['process_name'].tolist()
        _proc_list    = list(zip(proc_df['process_code'].tolist(),
                                 proc_df['process_name'].tolist()))
    else:
        process_names = []
        _proc_list    = []

    # Reverse lookup: safe field prefix → process_code (for ship_date/nbd edits on regular rows)
    _safe_to_pc = {_make_proc_field(pn): pc for pc, pn in _proc_list}

    # ── [CP-1] Expected column fields from vendor_processes ──────────────────
    _sched_log('[CP-1] process_names from vendor_processes:')
    for pc, pn in _proc_list:
        safe = _make_proc_field(pn)
        _sched_log(f'  pc={pc!r}  pn={pn!r}  safe={safe!r}')

    # Pivot to wide format
    row_data = _pivot_schedule(sched_df) if not sched_df.empty else []

    # ── [CP-2] Full pivot result — one line per (lot, process) ───────────────
    expected_safes = [_make_proc_field(pn) for _, pn in _proc_list]
    _sched_log(f'[CP-2] pivot result ({len(row_data)} rows):')
    for r in row_data:
        lot_id = r.get('lot_id')
        process_fields = {s: r.get(f'{s}_ship_date', '<KEY_MISSING>') for s in expected_safes}
        _sched_log(f'  lot={lot_id!r}  rk={r.get("_row_key")!r}  ' + '  '.join(f'{s}={v!r}' for s, v in process_fields.items()))

    # Merge SQLite comments into rows
    def _merge_comments(rdata: list, vc: str) -> None:
        # Collect every schedule_id from all pivoted rows (Oracle IDs are hex strings)
        all_sids = []
        for r in rdata:
            for k, v in r.items():
                if k.endswith('_schedule_id') and v:
                    all_sids.append(str(v))

        cmap_sid = _DO.get_vendor_sched_comments(vc, all_sids)   # {schedule_id: comment}
        # Fallback: original Oracle lot_id lookup covers rows saved with a placeholder schedule_id
        all_orig_lots = [r.get('_orig_lot_id') or r.get('lot_id', '') for r in rdata]
        all_orig_lots = [x for x in all_orig_lots if x]
        cmap_lot = _DO.get_vendor_sched_comments_by_lot(vc, all_orig_lots) if all_orig_lots else {}

        for r in rdata:
            r.setdefault('_comment_sid', '')
            # Primary: find comment by first matching Oracle schedule_id
            c = None
            matched_sid = None
            for k, v in r.items():
                if k.endswith('_schedule_id') and v:
                    sid = str(v)
                    c = cmap_sid.get(sid)
                    if c:
                        matched_sid = sid
                        break
            # Fallback: look up by original Oracle lot_id
            if c is None:
                orig = r.get('_orig_lot_id') or r.get('lot_id', '')
                c = cmap_lot.get(orig)
                if c:
                    matched_sid = str(c.schedule_id)
            if c:
                r['_comment_sid'] = matched_sid
                # SQLite db_drawing/product_name override Oracle only when Oracle returned empty
                if c.db_drawing:
                    r['db_drawing']   = c.db_drawing
                if c.product_name:
                    r['product_name'] = c.product_name
                r['cust_pn']  = c.cust_pn  or ''
                r['quantity'] = c.quantity  if c.quantity is not None else 1
                r['comment1'] = c.comment1  or ''
                r['comment2'] = c.comment2  or ''
                r['comment3'] = c.comment3  or ''

    _merge_comments(row_data, vendor_code)

    # ── [CP-3] Data sent to grid after merge_comments ────────────────────────
    _sched_log(f'[CP-3] row_data after _merge_comments ({len(row_data)} rows):')
    for r in row_data:
        proc_vals = {k: v for k, v in r.items() if k.endswith('_ship_date') or k.endswith('_schedule_id')}
        _sched_log(f'  lot={r.get("lot_id")!r}  rk={r.get("_row_key")!r}  sid={r.get("_comment_sid")!r}  {proc_vals}')

    # Build cust_pn suggestion map: {db_drawing: [cust_pn, most-recent-first]}
    _cust_pn_map: dict = {}

    def _build_cust_pn_map(rdata: list) -> None:
        _cust_pn_map.clear()
        for r in rdata:
            db = r.get('db_drawing', '')
            cp = r.get('cust_pn', '')
            if db and cp and cp not in _cust_pn_map.get(db, []):
                _cust_pn_map.setdefault(db, []).append(cp)

    _build_cust_pn_map(row_data)

    col_defs = _generate_vendor_sched_col_defs(process_names)

    # Always store latest params so the refresh button can re-query
    _sched_state['params']       = params
    _sched_state['vendor_code']  = vendor_code
    _sched_state['vendor_name']  = vendor_name

    # If tab already open for this vendor, refresh in place and switch to it
    existing_tab_id = _sched_vendor_tabs.get(vendor_code)
    if existing_tab_id is not None and existing_tab_id in current_tabs and _sched_state.get('grid') is not None:
        if _sched_state['toggle_form']:
            await _sched_state['toggle_form'](False)
        await _sched_state['grid'].run_grid_method('setGridOption', 'rowData', row_data)
        await _sched_state['grid'].run_grid_method('setGridOption', 'columnDefs', col_defs)
        tabs.set_value(current_tabs[existing_tab_id][1])
        spinner_close()
        return

    # First-creation path — only reached when no tab exists for this vendor.
    # The date-filter dialog (_date_filter_dlg) is created here once;
    # its closure variables remain valid for the tab's lifetime because
    # _on_sched_tab_delete removes the vendor from _sched_vendor_tabs, ensuring
    # this path re-runs (with fresh state) if the tab is closed and reopened.
    with tabs:
        tab_label = f'日程表_{vendor_name}'
        new_tab = ui.tab(tab_label).classes('rounded-t-lg bg-purple-200').props('closable')
        with new_tab:
            with ui.context_menu():
                def _on_sched_tab_delete(_vc=vendor_code, _tab=new_tab):
                    _sched_vendor_tabs.pop(_vc, None)
                    _sched_state_map.pop(_vc, None)
                    tab_delete(_tab, _tab.id)
                ui.menu_item('削除', on_click=_on_sched_tab_delete)
    with tab_panels:
        with CustomTabPanel(new_tab).classes('w-full h-full') as panel:

            # ── Shared state for add-row feature ───────────────────────────
            _refs       : dict = {}   # UI element refs (grid, buttons)
            _pinned_row : dict = {}   # live mirror of the pinned top row

            _ctx_event:       dict = {}   # full cellContextMenu event args (includes rowIndex)
            _ctx_row:         dict = {}   # right-clicked row data
            _editing_row_key: list = [None]  # [0] = _row_key of the row currently in edit mode
            _edit_overrides:  dict = {}   # buffered db_drawing/product_name changes during edit mode

            _LOT_RE = re.compile(r'^(?:42\d{8}|[KNS](?:\d|-){6,10})?$')  # empty OR real lot

            def _new_empty_row() -> dict:
                row: dict = {
                    'lot_id': '', '_orig_lot_id': '', 'db_drawing': '', 'product_name': '',
                    'cust_pn': '', 'quantity': 1,
                    'comment1': '', 'comment2': '', 'comment3': '',
                    'lot_status': '', '_comment_sid': '', '_row_key': '_new_',
                }
                for pn in process_names:
                    safe = _make_proc_field(pn)
                    row[f'{safe}_ship_date']          = ''
                    row[f'{safe}_nbd']                = ''
                    row[f'{safe}_vendor_commit_date'] = ''
                    row[f'{safe}_schedule_id']        = ''
                return row

            async def _toggle_add_row(show: bool) -> None:
                if show:
                    _pinned_row.clear()
                    _pinned_row.update(_new_empty_row())
                    await _refs['grid'].run_grid_method(
                        'setGridOption', 'pinnedTopRowData', [_pinned_row])
                    _refs['add_btn'].set_visibility(False)
                    _refs['save_btn'].set_visibility(True)
                    _refs['cancel_btn'].set_visibility(True)
                else:
                    await _refs['grid'].run_grid_method(
                        'setGridOption', 'pinnedTopRowData', [])
                    _pinned_row.clear()
                    _refs['add_btn'].set_visibility(True)
                    _refs['save_btn'].set_visibility(False)
                    _refs['cancel_btn'].set_visibility(False)

            # Store cancel handle so refresh-in-place can close an open form
            _sched_state['toggle_form'] = _toggle_add_row

            async def _do_schedule_grid_refresh() -> None:
                """Full grid refresh for use as on_save callback. Does not touch SQLite."""
                vc = _sched_state['vendor_code']
                new_sched_df = await M.fast_query('vendor_schedule', params, True)
                if new_sched_df is None:
                    return
                new_sched_df.columns = new_sched_df.columns.str.lower()
                if not new_sched_df.empty and 'schedule_id' in new_sched_df.columns:
                    new_sched_df = new_sched_df.drop_duplicates(
                        subset=['lot_id', 'process_name', 'schedule_id'])
                new_row_data = _pivot_schedule(new_sched_df) if not new_sched_df.empty else []
                _merge_comments(new_row_data, vc)
                await _refs['grid'].run_grid_method('setGridOption', 'rowData', new_row_data)

            async def _save_new_row() -> None:
                # Commit any active cell edit in the browser first, then yield so
                # _on_cell_changed can update _pinned_row before we read it.
                await grid.run_grid_method('stopEditing')
                await _asyncio.sleep(0.2)

                vc = _sched_state['vendor_code']

                # --- lot_id: empty (→ Oracle assigns; schedule_id used as key) or real lot ---
                lot_id_raw = (_pinned_row.get('lot_id') or '').strip().upper()
                if not _LOT_RE.match(lot_id_raw):
                    ui.notification(
                        'ロットIDの形式が正しくありません（空欄、または42＋8桁 / K・N・S始まりの形式）',
                        type='warning')
                    return

                # --- at least one of db_drawing / product_name / cust_pn ---
                db_drawing   = (_pinned_row.get('db_drawing')   or '').strip()[:200]
                product_name = (_pinned_row.get('product_name') or '').strip()[:200]
                cust_pn      = (_pinned_row.get('cust_pn')      or '').strip()[:100]
                if not db_drawing and not product_name and not cust_pn:
                    ui.notification(
                        'DB図番、品名、客先品番のいずれかを入力してください',
                        type='warning') #REFACTOR: Replace with existing custom notification function
                    return

                # --- quantity: integer >= 1 ---
                try:
                    qty = int(_pinned_row.get('quantity') or 1)
                    if qty < 1:
                        raise ValueError
                except (ValueError, TypeError):
                    ui.notification('数量は1以上の整数を入力してください', type='warning') #REFACTOR: Replace with existing custom notification function
                    return

                # --- comments (strip & cap) ---
                comment1 = (_pinned_row.get('comment1') or '').strip()[:500]
                comment2 = (_pinned_row.get('comment2') or '').strip()[:500]
                comment3 = (_pinned_row.get('comment3') or '').strip()[:500]

                # --- process dates: collect non-empty; ship_date required per process ---
                inserts = []
                for pc, pn in _proc_list:
                    safe    = _make_proc_field(pn)
                    sd_raw  = (_pinned_row.get(f'{safe}_ship_date')          or '').strip()
                    nd_raw  = (_pinned_row.get(f'{safe}_nbd')                or '').strip()
                    vcd_raw = (_pinned_row.get(f'{safe}_vendor_commit_date') or '').strip()
                    if not sd_raw and not nd_raw and not vcd_raw:
                        continue
                    if not sd_raw:
                        ui.notification(f'工程「{pn}」の支給日を入力してください', type='warning') #REFACTOR: Replace with existing custom notification function
                        return
                    sd = _normalize_date(sd_raw)
                    if sd is None:
                        ui.notification(
                            f'支給日の日付形式を認識できません: {sd_raw!r}', type='warning') #REFACTOR: Replace with existing custom notification function
                        return
                    nd = _normalize_date(nd_raw) if nd_raw else None
                    if nd_raw and nd is None:
                        ui.notification(
                            f'納期の日付形式を認識できません: {nd_raw!r}', type='warning') #REFACTOR: Replace with existing custom notification function
                        return
                    vcd = _normalize_date(vcd_raw) if vcd_raw else None
                    if vcd_raw and vcd is None:
                        ui.notification(
                            f'回答納期の日付形式を認識できません: {vcd_raw!r}', type='warning') #REFACTOR: Replace with existing custom notification function
                        return
                    inserts.append({
                        'process_code':       pc,
                        'safe':               safe,
                        'ship_date':          sd,   # normalized YYYY-MM-DD
                        'nbd':                nd,   # normalized YYYY-MM-DD or None
                        'vendor_commit_date': vcd,  # normalized YYYY-MM-DD or None
                    })

                if not inserts:
                    ui.notification('少なくとも1つの工程に支給日を入力してください', type='warning') #REFACTOR: Replace with existing custom notification function
                    return

                # --- Oracle MERGE per process ---
                # Oracle treats '' as NULL; NULL = NULL is never TRUE in a MERGE ON
                # condition, so an empty lot_id would cause all MERGEs to fail silently.
                # Generate a unique placeholder when the user left lot_id blank.
                import time as _time
                if not lot_id_raw:
                    oracle_lot_id = f'-T{int(_time.time())}'
                    _sched_log(f'[SAVE] empty lot_id → generated placeholder: {oracle_lot_id!r}')
                else:
                    oracle_lot_id = lot_id_raw

                _sched_log(f'[SAVE] oracle_lot_id={oracle_lot_id!r}  vc={vc!r}  inserts={inserts}')
                for ins in inserts:
                    _sched_log(f'  [SAVE] MERGE lot={oracle_lot_id!r} proc={ins["process_code"]!r} ship={ins["ship_date"]!r} nbd={ins["nbd"]!r}')
                    await M.fast_query(
                        'vendor_schedule_insert',
                        {'lot_id':       oracle_lot_id,
                         'vendor_code':  vc,
                         'process_code': ins['process_code'],
                         'ship_date':    ins['ship_date'],
                         'nbd':          ins['nbd']},
                        False, commit=True,
                    )
                    _sched_log(f'  [SAVE] MERGE done')
                # Build lookup: safe_prefix → vendor_commit_date for post-insert update
                _vcd_by_safe = {ins['safe']: ins['vendor_commit_date']
                                for ins in inserts if ins.get('vendor_commit_date')}

                # --- Refresh grid (same logic as initial load: wide-range + dedup) ---
                new_sched_df = await M.fast_query('vendor_schedule', params, True)
                _sched_log(f'[REFRESH] query returned: {None if new_sched_df is None else len(new_sched_df)} rows')
                if new_sched_df is not None:
                    new_sched_df.columns = new_sched_df.columns.str.lower()
                    if not new_sched_df.empty and 'schedule_id' in new_sched_df.columns:
                        new_sched_df = new_sched_df.drop_duplicates(
                            subset=['lot_id', 'process_name', 'schedule_id'])
                    new_row_data = _pivot_schedule(new_sched_df) if not new_sched_df.empty else []
                    _sched_log(f'[REFRESH] new_row_data={len(new_row_data)} grid rows  oracle_lot_present={any(r.get("_orig_lot_id") == oracle_lot_id for r in new_row_data)}')

                    # Pick the first schedule_id of the newly saved lot to own the SQLite comment row.
                    # Match by _orig_lot_id (= oracle_lot_id) because lot_id in the grid is '-'
                    # for all non-real lots.
                    lot_first_sid = None
                    new_lot_row   = None
                    _sched_log(f'[SAVE] searching new_row_data ({len(new_row_data)} rows) for oracle_lot_id={oracle_lot_id!r}')
                    for r in new_row_data:
                        orig = r.get('_orig_lot_id', r.get('lot_id', ''))
                        _sched_log(f'  [SAVE] row orig_lot={orig!r} lot_id={r.get("lot_id")!r}')
                        if orig.upper() == oracle_lot_id.upper():
                            new_lot_row = r
                            for k, v in r.items():
                                if k.endswith('_schedule_id') and v:
                                    lot_first_sid = str(v)
                                    break
                            break
                    _sched_log(f'[SAVE] lot_first_sid={lot_first_sid!r}')

                    # --- Update vendor_commit_date for each process that had one ---
                    if _vcd_by_safe and new_lot_row:
                        for safe_prefix, vcd in _vcd_by_safe.items():
                            sid = new_lot_row.get(f'{safe_prefix}_schedule_id')
                            if sid:
                                _sched_log(f'[SAVE] vendor_commit_update sid={sid!r} vcd={vcd!r}')
                                await M.fast_query('vendor_commit_update',
                                                   {'commit_date': vcd, 'schedule_id': sid},
                                                   False, commit=True)

                    # Fallback: negative timestamp string if Oracle didn't return this lot
                    if not lot_first_sid:
                        lot_first_sid = str(-int(_time.time()))
                        _sched_log(f'[SAVE] fallback lot_first_sid={lot_first_sid!r}')

                    # --- SQLite upsert ---
                    _sched_log(f'[SQLITE] upsert sid={lot_first_sid!r} lot={oracle_lot_id!r} vc={vc!r} db={db_drawing!r} pn={product_name!r} cpn={cust_pn!r}')
                    _sqlite_result = _DO.upsert_vendor_sched_comment(
                        lot_first_sid, oracle_lot_id, vc,
                        db_drawing   = db_drawing    or None,
                        product_name = product_name  or None,
                        cust_pn      = cust_pn       or None,
                        quantity     = qty,
                        comment1     = comment1      or None,
                        comment2     = comment2      or None,
                        comment3     = comment3      or None,
                    )
                    _sched_log(f'[SQLITE] upsert result={_sqlite_result!r}')

                    _merge_comments(new_row_data, vc)
                    await _refs['grid'].run_grid_method('setGridOption', 'rowData', new_row_data)
                    # Update cust_pn suggestions from refreshed data
                    if cust_pn and db_drawing:
                        _update_cust_pn_map(db_drawing, cust_pn)
                    _build_cust_pn_map(new_row_data)
                    ui.run_javascript(f'window._custPnMap = {json.dumps(_cust_pn_map)}')

                ui.notification('登録しました', type='positive')
                await _toggle_add_row(False)

            # Async wrappers — NiceGUI schedules these via background_tasks
            async def _on_add_clicked(_=None):
                await _toggle_add_row(True)

            async def _on_cancel_clicked(_=None):
                if _editing_row_key[0] is not None:
                    # Exit edit mode: revert grid, restore col defs, hide buttons
                    _editing_row_key[0] = None
                    _edit_overrides.clear()
                    await grid.run_grid_method('stopEditing', True)  # True = cancel/revert
                    await grid.run_grid_method(
                        'setGridOption', 'columnDefs', _generate_vendor_sched_col_defs(process_names))
                    _refs['add_btn'].set_visibility(True)
                    _refs['save_btn'].set_visibility(False)
                    _refs['cancel_btn'].set_visibility(False)
                else:
                    await _toggle_add_row(False)

            async def _save_edit_row():
                await grid.run_grid_method('stopEditing')  # commit active cell edit
                sid = _edit_overrides.get('_comment_sid')
                if not sid:
                    for k, v in _ctx_row.items():
                        if k.endswith('_schedule_id') and v:
                            sid = str(v)
                            break
                if sid:
                    vc  = _sched_state['vendor_code']
                    lot = _edit_overrides.get('lot_id', '')
                    _DO.upsert_vendor_sched_comment(
                        str(sid), lot, vc,
                        db_drawing   = _edit_overrides.get('db_drawing')    or None,
                        product_name = _edit_overrides.get('product_name')  or None,
                    )
                    ui.notification('保存しました', type='positive')
                # Restore col defs + button state
                _editing_row_key[0] = None
                _edit_overrides.clear()
                await grid.run_grid_method(
                    'setGridOption', 'columnDefs', _generate_vendor_sched_col_defs(process_names))
                _refs['add_btn'].set_visibility(True)
                _refs['save_btn'].set_visibility(False)
                _refs['cancel_btn'].set_visibility(False)
                await _on_refresh_clicked()

            async def _on_save_clicked(_=None):
                if _editing_row_key[0] is not None:
                    await _save_edit_row()
                else:
                    await _save_new_row()

            async def _on_refresh_clicked(_=None):
                stored = _sched_state.get('params')
                if stored:
                    await _open_supplier_schedule_sample(stored)

            # ── Individual date-picker sub-dialogs ──────────────────────────
            with ui.dialog() as _from_picker_dlg, ui.card().classes('p-2'):
                _from_picker = ui.date().props('mask="YYYY/MM/DD"')
                def _confirm_from(_=None):
                    if _from_picker.value:
                        _from_btn.set_text('開始日：' + _from_picker.value)
                    _from_picker_dlg.close()
                ui.button('確定', on_click=_confirm_from).props('color=primary flat')

            with ui.dialog() as _to_picker_dlg, ui.card().classes('p-2'):
                _to_picker = ui.date().props('mask="YYYY/MM/DD"')
                def _confirm_to(_=None):
                    if _to_picker.value:
                        _to_btn.set_text('終了日：' + _to_picker.value)
                    _to_picker_dlg.close()
                ui.button('確定', on_click=_confirm_to).props('color=primary flat')

            # ── Date filter dialog (created once; reused on every open) ────
            # Not stored in _refs — never needs to be hidden programmatically.
            with ui.dialog() as _date_filter_dlg, ui.card().classes('p-4 gap-3'):
                ui.label('期間フィルター').classes('font-bold text-base')
                with ui.row().classes('gap-2 items-center'):
                    _from_btn = (
                        ui.button('開始日：--/--/--')
                        .props('outline color=primary')
                        .on('click', _from_picker_dlg.open)
                    )
                    ui.label('〜')
                    _to_btn = (
                        ui.button('終了日：--/--/--')
                        .props('outline color=primary')
                        .on('click', _to_picker_dlg.open)
                    )

                async def _on_apply_filter(_=None):
                    raw_from = (_from_picker.value or '').replace('/', '-')
                    raw_to   = (_to_picker.value   or '').replace('/', '-')
                    if not raw_from or not raw_to:
                        ui.notification('開始日と終了日を選択してください', type='warning')
                        return
                    if raw_from > raw_to:  # ISO format: lexicographic == chronological
                        ui.notification('開始日は終了日以前を設定してください', type='warning')
                        return
                    new_params = dict(_sched_state['params'])
                    new_params['date_from'] = raw_from
                    new_params['date_to']   = raw_to
                    _date_filter_dlg.close()
                    _sched_state['params'] = new_params
                    await _open_supplier_schedule_sample(new_params)

                with ui.row().classes('justify-end gap-2 mt-2'):
                    ui.button('キャンセル', on_click=_date_filter_dlg.close).props('flat')
                    ui.button('適用', on_click=_on_apply_filter).props('color=primary')

            def _on_filter_clicked(_=None):
                p = _sched_state.get('params') or {}
                from_str = (p.get('date_from') or '').replace('-', '/')
                to_str   = (p.get('date_to')   or '').replace('-', '/')
                _from_picker.value = from_str
                _to_picker.value   = to_str
                _from_btn.set_text('開始日：' + from_str if from_str else '開始日：未設定')
                _to_btn.set_text('終了日：' + to_str if to_str else '終了日：未設定')
                _date_filter_dlg.open()

            _split_active: list = [False]  # [0] = True when right panel is visible

            with ui.card().classes('w-full h-full flex flex-col bg-purple-50') as card:
                # ── Title bar ──────────────────────────────────────────────────
                with ui.row().classes('w-full items-center mb-1 gap-1'):
                    ui.label('仕入先スケジュール表').classes('text-lg font-bold')
                    ui.space()
                    (
                        ui.button(icon='calendar_month')
                        .props('flat round color=grey')
                        .tooltip('期間を変更')
                        .on('click', _on_filter_clicked)
                    )
                    (
                        ui.button(icon='refresh')
                        .props('flat round color=grey')
                        .tooltip('データを更新')
                        .on('click', _on_refresh_clicked)
                    )
                    _refs['add_btn'] = (
                        ui.button(icon='add')
                        .props('flat round color=primary')
                        .tooltip('新規行を追加')
                        .on('click', _on_add_clicked)
                    )
                    _refs['save_btn'] = (
                        ui.button(icon='save')
                        .props('flat round color=positive')
                        .tooltip('保存')
                        .on('click', _on_save_clicked)
                    )
                    _refs['save_btn'].set_visibility(False)
                    _refs['cancel_btn'] = (
                        ui.button(icon='close')
                        .props('flat round dense')
                        .tooltip('キャンセル')
                        .on('click', _on_cancel_clicked)
                    )
                    _refs['cancel_btn'].set_visibility(False)
                    split_btn = (
                        ui.button(icon='call_split')
                        .props('flat round color=grey')
                        .tooltip('分割ビュー')
                    )

                # ── Split layout: one shared card; right panel snaps closed past threshold ──
                with ui.card().classes('w-full h-full') as split_card:
                    with ui.splitter(value=100).classes('w-full h-full') as spl:
                        with spl.separator:
                            ui.icon('drag_indicator', color='primary').classes('cursor-move text-4xl')
                        with spl.before:
                            with ui.element('div').classes('w-full h-full') as _ag_container:
                                grid = ui.aggrid({
                                    'columnDefs': col_defs,
                                    'rowData': row_data,
                                    'defaultColDef': {'resizable': True, 'sortable': True},
                                    'singleClickEdit': False,
                                    'stopEditingWhenCellsLoseFocus': True,
                                    'rowSelection': 'multiple',
                                    ':getRowId': '(params) => params.data._row_key || params.data.lot_id',
                                    ':getRowStyle': "(params) => params.node.rowPinned === 'top' ? {'background': '#fffde7'} : null",
                                }, auto_size_columns=False, theme='balham').classes('w-full').style('height: calc(100vh - 160px)')
                        with spl.after:
                            after_wrapper = ui.element('div').classes('w-full h-full')

                _refs['grid']         = grid
                _sched_state['grid']  = grid

                # Inject cust_pn suggestion map into the browser

                def _update_cust_pn_map(db_drawing: str, cust_pn: str) -> None:
                    if not db_drawing or not cust_pn:
                        return
                    lst = _cust_pn_map.setdefault(db_drawing, [])
                    if cust_pn in lst:
                        lst.remove(cust_pn)
                    lst.insert(0, cust_pn)
                    ui.run_javascript(f'window._custPnMap = {json.dumps(_cust_pn_map)}')

                ui.run_javascript(f'window._custPnMap = {json.dumps(_cust_pn_map)}')
                # Initialise lot_id-editing guard flag
                ui.run_javascript('window._lotIdEditing = false')

                # ── lot_id editing guard — disable auto-fill targets while lot_id is active ──
                def _on_cell_editing_started(e):
                    if e.args.get('colId') == 'lot_id' and e.args.get('rowPinned') == 'top':
                        ui.run_javascript('window._lotIdEditing = true')

                def _on_cell_editing_stopped(e):
                    if e.args.get('colId') == 'lot_id':
                        ui.run_javascript('window._lotIdEditing = false')

                grid.on('cellEditingStarted', _on_cell_editing_started)
                grid.on('cellEditingStopped', _on_cell_editing_stopped)

                # ── Cell edit handler ──────────────────────────────────────────
                _SQLITE_FIELDS = {'cust_pn', 'quantity', 'comment1', 'comment2', 'comment3'}

                spinner_close()

                async def _on_cell_changed(e):
                    col        = e.args.get('colId', '')
                    new_val    = e.args.get('newValue')
                    old_val    = e.args.get('oldValue')
                    row        = e.args.get('data', {})
                    is_pinned  = e.args.get('rowPinned') == 'top'
                    vc         = _sched_state['vendor_code']

                    # For regular rows skip if the value didn't actually change
                    if not is_pinned and new_val == old_val:
                        return

                    # ── Pinned (new) row ───────────────────────────────────────
                    if is_pinned:
                        # Sanitise and store in our Python-side mirror
                        if col == 'quantity':
                            try:
                                _pinned_row[col] = int(new_val) if new_val not in (None, '') else 1
                            except (ValueError, TypeError):
                                _pinned_row[col] = 1
                        elif col in ('lot_id', 'db_drawing', 'product_name',
                                     'cust_pn', 'comment1', 'comment2', 'comment3'):
                            # Strip, cap at 500 chars, keep printable
                            _pinned_row[col] = (str(new_val or '')).strip()[:500]
                        elif col.endswith('_ship_date') or col.endswith('_nbd') or col.endswith('_vendor_commit_date'):
                            raw = (str(new_val or '')).strip()
                            _pinned_row[col] = _normalize_date(raw) or raw
                        else:
                            _pinned_row[col] = new_val

                        # Auto-fill db_drawing / product_name / lot_status when lot_id is set
                        if col == 'lot_id' and new_val:
                            lot_id_clean = (str(new_val)).strip().upper()
                            if _LOT_RE.match(lot_id_clean):
                                df = await M.fast_query(
                                    'lot_info', {'lot_id': lot_id_clean}, True)
                                if df is not None and not df.empty:
                                    df.columns = df.columns.str.lower()
                                    r0 = df.iloc[0]
                                    _pinned_row['db_drawing']   = _safe_str(r0.get('db_drawing'))
                                    _pinned_row['product_name'] = _safe_str(r0.get('product_name'))
                                    _pinned_row['lot_status']   = _safe_str(r0.get('lot_status'))
                                    # Suggest cust_pn from map if not already set
                                    db = _pinned_row.get('db_drawing', '')
                                    suggestions = _cust_pn_map.get(db, [])
                                    if suggestions and not _pinned_row.get('cust_pn'):
                                        _pinned_row['cust_pn'] = suggestions[0]
                                    # Push updated values back to the pinned row in the grid
                                    await grid.run_grid_method(
                                        'setGridOption', 'pinnedTopRowData', [dict(_pinned_row)])
                        return

                    # ── Regular row ───────────────────────────────────────────
                    lot = row.get('lot_id', '')
                    if not lot or not vc:
                        return
                    saved = False
                    if col in _SQLITE_FIELDS:
                        val = int(new_val) if col == 'quantity' and new_val not in (None, '') else new_val
                        # Determine which schedule_id owns (or will own) this row's SQLite comment
                        sid = row.get('_comment_sid')
                        if not sid:
                            # No comment row yet — pick first process schedule_id to create one
                            for k, v in row.items():
                                if k.endswith('_schedule_id') and v:
                                    sid = str(v)
                                    break
                        if sid:
                            _DO.upsert_vendor_sched_comment(str(sid), lot, vc, **{col: val})
                            saved = True
                        if col == 'cust_pn' and new_val:
                            _update_cust_pn_map(row.get('db_drawing', ''), str(new_val))
                    elif col in ('db_drawing', 'product_name'):
                        if _editing_row_key[0] and row.get('_row_key') == _editing_row_key[0]:
                            _edit_overrides[col] = (str(new_val or '')).strip()[:500]
                            # Do NOT save immediately — wait for 保存 button
                    elif col.endswith('_vendor_commit_date'):
                        safe_proc = col[:-len('_vendor_commit_date')]
                        sid = row.get(f'{safe_proc}_schedule_id')
                        if sid:
                            await M.fast_query('vendor_commit_update',
                                               {'commit_date': new_val or '', 'schedule_id': sid},
                                               False, commit=True)
                            saved = True
                    elif col.endswith('_ship_date'):
                        if not new_val:
                            ui.notification('支給日は必須です', type='warning')
                            return
                        sd = _normalize_date(str(new_val).strip())
                        if sd is None:
                            ui.notification(f'日付形式を認識できません: {new_val!r}', type='warning')
                            return
                        safe_proc = col[:-len('_ship_date')]
                        pc = _safe_to_pc.get(safe_proc)
                        if pc:
                            nbd_val = row.get(f'{safe_proc}_nbd') or None
                            await M.fast_query('vendor_schedule_insert',
                                               {'lot_id': lot, 'vendor_code': vc, 'process_code': pc,
                                                'ship_date': sd, 'nbd': nbd_val},
                                               False, commit=True)
                            saved = True
                    elif col.endswith('_nbd'):
                        nd = _normalize_date(str(new_val).strip()) if new_val else None
                        if new_val and nd is None:
                            ui.notification(f'日付形式を認識できません: {new_val!r}', type='warning')
                            return
                        safe_proc = col[:-len('_nbd')]
                        pc = _safe_to_pc.get(safe_proc)
                        if pc:
                            sd_val = row.get(f'{safe_proc}_ship_date') or None
                            if not sd_val:
                                ui.notification('先に支給日を入力してください', type='warning')
                                return
                            await M.fast_query('vendor_schedule_insert',
                                               {'lot_id': lot, 'vendor_code': vc, 'process_code': pc,
                                                'ship_date': sd_val, 'nbd': nd},
                                               False, commit=True)
                            saved = True
                    if saved:
                        ui.notification('保存しました', type='positive')

                grid.on('cellValueChanged', _on_cell_changed)

                # ── Context menu: right-click cache ───────────────────────────────
                def _on_cell_context_menu(e):
                    _ctx_event.clear()
                    _ctx_event.update(e.args)
                    _ctx_row.clear()
                    _ctx_row.update(e.args.get('data', {}))

                # ── 行削除 ────────────────────────────────────────────────────────
                _deleting: list = [False]  # [0] = guard against concurrent delete calls

                async def _on_delete_rows_clicked(_=None):
                    if _deleting[0]:
                        return
                    selected = await grid.get_selected_rows()
                    rows_to_delete = selected if selected else ([dict(_ctx_row)] if _ctx_row else [])
                    if not rows_to_delete:
                        return
                    _deleting[0] = True
                    spinner_open()
                    try:
                        for row in rows_to_delete:
                            sids = [v for k, v in row.items() if k.endswith('_schedule_id') and v]
                            for sid in sids:
                                await M.fast_query('vendor_schedule_delete', {'schedule_id': sid},
                                                   False, commit=True)
                            comment_sid = row.get('_comment_sid')
                            if comment_sid:
                                _DO.delete_vendor_sched_comment(str(comment_sid))
                        ui.notification(f'{len(rows_to_delete)}行削除しました', type='positive')
                        await _on_refresh_clicked()
                    finally:
                        _deleting[0] = False
                        spinner_close()

                # ── 編集: enter in-place edit mode ────────────────────────────────
                async def _on_edit_row_clicked(_=None):
                    if not _ctx_row:
                        return
                    # Don't allow edit mode if add-row mode is already active
                    if _refs['save_btn'].visible:
                        return
                    row_key   = _ctx_row.get('_row_key', '')
                    row_index = _ctx_event.get('rowIndex')
                    if not row_key or row_index is None:
                        return
                    grid.run_grid_method('deselectAll')
                    # Buffer initial values so save knows the original state
                    _editing_row_key[0] = row_key
                    _edit_overrides.clear()
                    _edit_overrides.update({
                        'db_drawing':   (_ctx_row.get('db_drawing')   or '').strip(),
                        'product_name': (_ctx_row.get('product_name') or '').strip(),
                        '_comment_sid': _ctx_row.get('_comment_sid') or '',
                        'lot_id':       _ctx_row.get('lot_id') or '',
                    })

                    # Temporarily make db_drawing + product_name editable for this row only
                    editable_for_row = (
                        f'(params) => (params.node.rowPinned === "top" && !window._lotIdEditing)'
                        f' || params.data._row_key === "{row_key}"'
                    )
                    new_col_defs = _generate_vendor_sched_col_defs(process_names)
                    for cd in new_col_defs:
                        if cd.get('field') in ('db_drawing', 'product_name'):
                            cd[':editable'] = editable_for_row
                    await grid.run_grid_method('setGridOption', 'columnDefs', new_col_defs)

                    # Start editing the db_drawing cell in-place
                    await grid.run_grid_method(
                        'startEditingCell', {'rowIndex': row_index, 'colKey': 'db_drawing'})

                    # Show 保存 / キャンセル buttons (reuse existing title-bar buttons)
                    _refs['add_btn'].set_visibility(False)
                    _refs['save_btn'].set_visibility(True)
                    _refs['cancel_btn'].set_visibility(True)

            with _ag_container:
                with ui.context_menu():
                    ui.menu_item('行削除', on_click=_on_delete_rows_clicked)
                    ui.separator()
                    ui.menu_item('編集',   on_click=_on_edit_row_clicked)

                grid.on('cellContextMenu', _on_cell_context_menu)

            def _build_right_panel(container):
                _shipping_ctx_row: dict = {}
                import asyncio
                _right_current_tabs: dict = {}  # name → (tab, panel)

                with container:
                    # ── Initial controls (visible until first search) ──────
                    with ui.element('div').classes('p-4 w-full') as initial_view:
                        ui.toggle(
                            options={'i': '指図進捗', 'w': '仕掛状況', 't': '関連指図'},
                            value='i',
                        ).classes('w-full justify-center mb-4').bind_value(search_mode).on_value_change(
                            lambda e: search_mode.__setitem__('lot', e.value in 'it')
                        )
                        right_textarea = ui.textarea(placeholder='検索').props('clearable').classes('w-full mb-2')
                        ui.button('検索', icon='search', on_click=lambda: asyncio.ensure_future(_right_search())).classes('w-full')
                        ui.separator().classes('my-3')
                        ui.button(
                            '外注支給日程表を表示',
                            icon='table_chart',
                            on_click=lambda: asyncio.ensure_future(_right_show_shipping()),
                        ).classes('w-full')

                    # ── Mini tab system (hidden until first search) ────────
                    with ui.element('div').classes('w-full h-full') as right_tabs_container:
                        right_tabs = ui.tabs().classes('w-full')
                        right_panels = ui.tab_panels(right_tabs).classes('w-full h-full')
                    right_tabs_container.set_visibility(False)

                # ── Nested helpers ────────────────────────────────────────

                def _right_tab_delete(tab_name: str):
                    tab, panel = _right_current_tabs.pop(tab_name, (None, None))
                    if tab:
                        tab.delete()
                    if panel:
                        panel.delete()
                    if not _right_current_tabs:
                        right_tabs_container.set_visibility(False)
                        initial_view.set_visibility(True)

                async def _right_search():
                    keyword = right_textarea.value
                    if not keyword:
                        return
                    mode = search_mode['value']
                    try:
                        df_dict, query_obj = await query_data_stdin(mode, keyword)
                    except Exception:
                        return  # query_data_stdin shows its own error notifications
                    if not df_dict:
                        return

                    initial_view.set_visibility(False)
                    right_tabs_container.set_visibility(True)

                    mode_str = {'i': '指図', 'w': '仕掛', 't': '関係'}
                    last_tab = None
                    for key, df in df_dict.items():
                        tab_name = f'{mode_str.get(mode, mode)}_{key}'
                        _query_for_tab = query_obj.copy([key]) if mode == 'w' else query_obj.copy()
                        with right_tabs:
                            tab = ui.tab(tab_name)
                            with tab:
                                with ui.context_menu():
                                    ui.menu_item('削除', on_click=lambda n=tab_name: _right_tab_delete(n))
                        with right_panels:
                            with CustomTabPanel(tab).classes('w-full h-full') as panel:
                                with ui.card().classes('w-full h-full flex flex-col') as result_card:
                                    _render_result_panel(
                                        result_card, df, mode, tab_name,
                                        refresh_clean_key=_query_for_tab,
                                        vendor_code=vendor_code,
                                        vendor_name=_sched_state.get('vendor_name'),
                                        on_save=_do_schedule_grid_refresh,
                                    )
                        _right_current_tabs[tab_name] = (tab, panel)
                        last_tab = tab
                    if last_tab:
                        right_tabs.set_value(last_tab)
                    spinner_close()

                async def _right_show_shipping():
                    spinner_open()
                    df = await M.fast_query('NOKEY_asp_supply_schedule', [], True)
                    
                    if df is None or df.empty:
                        notification('warning', '結果なし')
                        spinner_close()
                        return
                    df['index_id'] = range(len(df))  

                    initial_view.set_visibility(False)
                    right_tabs_container.set_visibility(True)
                    
                    tab_name = '外注支給日程表'
                    if tab_name in _right_current_tabs:
                        right_tabs.set_value(_right_current_tabs[tab_name][0])
                        spinner_close()
                        return
                    with right_tabs:
                        tab = ui.tab(tab_name)
                        with tab:
                            with ui.context_menu():
                                    ui.menu_item('削除', on_click=lambda n=tab_name: _right_tab_delete(n))
                    try:
                        with right_panels:
                            with CustomTabPanel(tab).classes('w-full h-full') as panel:
                                with ui.card().classes('w-full h-full'):
                                    ag = AgGrid()
                                    _, skeleton_list = ag.init_column_settings('asp')
                                    _shipping_ag_container, grid = ag.aggrid_creator(df, _table_view = True, row_id= 'index_id')
                                    grid.on('cellContextMenu', lambda e: (
                                        _shipping_ctx_row.clear(),
                                        _shipping_ctx_row.update(e.args.get('data', {}))
                                    ))
                                    #ui.aggrid({
                                    #    'columnDefs': [
                                    #        {'field': col, 'headerName': col,
                                    #        'resizable': True, 'sortable': True}
                                    #        for col in df.columns
                                    #    ],
                                    #    'rowData': df.to_dict('records'),
                                    #    'defaultColDef': {'resizable': True, 'sortable': True},
                                    #}, auto_size_columns=False, theme='balham').classes('w-full h-full')
                    except Exception as e:
                        notification('negative', f'タブの表示に失敗: {e}')
                        spinner_close()
                        return
                    _right_current_tabs[tab_name] = (tab, panel)
                    right_tabs.set_value(tab)

                    async def _on_add_to_schedule_shipping():
                        selected = await grid.get_selected_rows() #Fixed Manually
                        rows     = selected if selected else [_shipping_ctx_row]
                        #print(selected)
                        lot_ids_sel = [r['指図'] for r in rows if r.get('指図')]
                        if not lot_ids_sel:
                            ui.notify('指図番号が取得できません', type='warning')
                            return
                        # vendor_code=None: shipping grid is cross-vendor — user must select supplier
                        await _quick_add_to_schedule_dialog(lot_ids_sel, vendor_code, on_save=_do_schedule_grid_refresh)

                    with _shipping_ag_container:
                        with ui.context_menu():
                            ui.menu_item('支給日程表に追加', on_click=_on_add_to_schedule_shipping)

                    skeletons_up(skeleton_list, False)
                    spinner_close()

            _build_right_panel(after_wrapper)

            def _activate_split():
                _split_active[0] = True
                spl.value = 50
                split_btn.props('color=primary')

            def _deactivate_split():
                _split_active[0] = False
                spl.value = 100
                split_btn.props('color=grey')

            def _toggle_split():
                if not _split_active[0]:
                    _activate_split()
                else:
                    _deactivate_split()

            # Wire the split button AFTER _toggle_split is defined
            split_btn.on('click', _toggle_split)

    current_tabs[new_tab.id] = (None, new_tab, panel)
    _sched_vendor_tabs[vendor_code] = new_tab.id
    tabs.set_value(new_tab)
    search_prompt['prompt'] = False  # hide home tab — same as newtab()


async def _open_reschedule_batch_dialog(lot_id_flow_product_triples: list) -> None:
    """Open the interactive Gantt batch reschedule dialog.

    Delegates entirely to gantt_dialog.open_gantt_dialog().
    """
    await open_gantt_dialog(
        lot_id_flow_product_triples,
        username=USER_INS.user_name or USER_INS.uid,
        excel_config=RESCHEDULE_EXCEL_CONFIG,
        query_fn=query_data_stdin,
        spinner_open=spinner_open,
        spinner_close=spinner_close,
    )


# ─────────────────────────────────────────────────────────────────────────────
def _render_result_panel(
    container: 'ui.element',
    dataframe,
    mode_flag: str,
    name: str,
    refresh_clean_key: 'M.query_',
    db_scope: 'str | None' = None,
    timestamp: str = '',
    tab_panel: 'ui.element | None' = None,
    vendor_code=None,    # str | None — None = show supplier selector
    vendor_name=None,    # str | None — shown in dialog subtitle when vendor is fixed
    on_save=None,        # async callable | None — called after successful write
) -> None:
    """Render AgGrid + column menu + fast-filter into `container`.

    Extracted from newtab's card block + post-card event bindings.
    Called by newtab and by the right panel in _open_supplier_schedule_sample.

    Args:
        container:          The ui.card or ui.element to render into.
        dataframe:          The search result DataFrame (or list of tuples for 't' mode).
        mode_flag:          Single-char mode ('i', 'w', 't', etc.).
        name:               Tab/panel name string (used for labels and filter keys).
        refresh_clean_key:  M.query_ used by the update-data button. Must not be None â
                            passing None causes AttributeError at refresh_clean_key.key
                            inside the update_row_data closure.
        db_scope:           DB code string for WIP mode (None for other modes).
        timestamp:          Time string shown in the panel; '' is acceptable.
    """
    if mode_flag != 't':
        aggrid_instance = AgGrid()
        with ui.row().classes('w-full'):
            # Eager: set aggrid_col_setting so aggrid_creator() can use it.
            # Does NOT build any UI widgets.
            aggrid_instance.init_indexed_col_header(mode_flag)
            valid_columns = {col['field'] for col in aggrid_instance.aggrid_col_setting}
            # Stubs — populated by _lazy_build_menu_shell() on first menu open.
            check_row_container: ui.element | None = None
            col_ck_sk_list: list = []
            _menu_refs: dict = {}
            _menu_shell_built = False
            _menu_gen: int = 0
            with ui.button(icon='menu').props('color=primary') as menu_btn:
                with CustomMenu().style('min-width:420px') as column_menu:
                    # ── Gradient header (static) ─────────────────────────
                    with ui.element('div').style(
                        'display:flex;align-items:center;gap:10px;padding:12px 16px;'
                        f'background:linear-gradient(135deg,{ACCENT_COLOR},#3d7fa3);width:100%;'
                    ):
                        ui.label('フィルター&表示設定').style('color:white;font-weight:600;font-size:14px;')

                    # ── Column visibility section ─────────────────────────
                    with ui.element('div').style(
                        'display:flex;align-items:center;gap:5px;'
                        'font-size:10px;font-weight:700;color:#94a3b8;'
                        'text-transform:uppercase;letter-spacing:.08em;padding:10px 16px 5px;'
                    ):
                        ui.html('<svg style="width:12px;height:12px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>', sanitize=False)
                        ui.label('列の表示/非表示設定')

                    with ui.element('div').style('padding:0 16px 8px;'):
                        _menu_refs['col_settings_slot'] = ui.element('div').classes('w-full')
                        with _menu_refs['col_settings_slot']:
                            with ui.element('div').classes('w-full flex flex-col gap-1') as _csp:
                                ui.skeleton('rect').classes('w-full h-8')
                                ui.skeleton('rect').classes('w-full h-8')
                                ui.skeleton('rect').classes('w-3/4 h-8')
                            _menu_refs['col_settings_spinner'] = _csp

                    with ui.element('div').style('display:flex;align-items:center;gap:8px;padding:4px 16px 12px;'):
                        _menu_refs['colapplybt'] = ui.button('適用').props('unelevated dense color=green').style(
                            'padding:5px 12px;font-size:12px;font-weight:600;border-radius:8px;'
                        )
                        _menu_refs['allselectbt'] = ui.button('全選択').style(
                            'padding:5px 12px;background:transparent;color:#64748b;'
                            'border:1px solid #e2e8f0;border-radius:7px;font-size:12px;font-weight:500;'
                        ).props('flat dense')
                        _menu_refs['allunselectbt'] = ui.button('全取消').style(
                            'padding:5px 12px;background:transparent;color:#64748b;'
                            'border:1px solid #e2e8f0;border-radius:7px;font-size:12px;font-weight:500;'
                        ).props('flat dense')

                    # Divider
                    ui.element('div').style('height:1px;background:#f0f4f8;margin:4px 12px;')

                    # ── Fast filter section ───────────────────────────
                    with ui.element('div').style(
                        'display:flex;align-items:center;gap:5px;'
                        'font-size:10px;font-weight:700;color:#94a3b8;'
                        'text-transform:uppercase;letter-spacing:.08em;padding:10px 16px 5px;'
                    ):
                        ui.html('<svg style="width:12px;height:12px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M3 4a1 1 0 011-1h16a1 1 0 011 1v2a1 1 0 01-.293.707L13 13.414V19a1 1 0 01-.553.894l-4 2A1 1 0 017 21v-7.586L3.293 6.707A1 1 0 013 6V4z"/></svg>', sanitize=False)
                        ui.label('ファストフィルター')

                    with ui.element('div').style('padding:0 16px 8px;display:flex;flex-direction:column;gap:6px;'):
                        # Built-in fast filter buttons
                        with ui.element('div').style(
                            'display:flex;align-items:center;gap:8px;width:100%;padding:8px 12px;'
                            'border-radius:10px;font-size:13px;font-weight:500;cursor:pointer;'
                            'background:#f1f5f9;color:#475569;border:1px solid #e2e8f0;'
                        ) as _preship:
                            ui.html('<svg style="width:14px;height:14px;color:#94a3b8;flex-shrink:0;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/></svg>', sanitize=False)
                            ui.label('梱包/支給前の指図').style('font-size:13px;')
                        _menu_refs['preship'] = _preship

                        with ui.element('div').style(
                            'display:flex;align-items:center;gap:8px;width:100%;padding:8px 12px;'
                            'border-radius:10px;font-size:13px;font-weight:500;cursor:pointer;'
                            'background:#f1f5f9;color:#475569;border:1px solid #e2e8f0;'
                        ) as _onhold:
                            ui.html('<svg style="width:14px;height:14px;color:#94a3b8;flex-shrink:0;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 9v6m4-6v6m7-3a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>', sanitize=False)
                            ui.label('保留中の指図').style('font-size:13px;')
                        _menu_refs['onhold'] = _onhold

                        with ui.element('div').style(
                            'display:flex;align-items:center;gap:8px;width:100%;padding:8px 12px;'
                            'border-radius:10px;font-size:13px;font-weight:500;cursor:pointer;'
                            'background:#f1f5f9;color:#475569;border:1px solid #e2e8f0;'
                        ) as _restricted:
                            ui.html('<svg style="width:14px;height:14px;color:#f59e0b;flex-shrink:0;" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"/></svg>', sanitize=False)
                            ui.label('黄袋履歴有指図').style('font-size:13px;')
                        _menu_refs['restricted'] = _restricted

                        # ── Date period filters (modes i & w) ──
                        if mode_flag in ('i', 'w'):
                            with ui.element('div').style('display:flex;align-items:center;gap:8px;margin-top:8px;'):
                                ui.element('div').style('flex:1;height:1px;background:#e8edf2;')
                                ui.label('最終更新日').style('font-size:10px;color:#94a3b8;font-weight:700;letter-spacing:.05em;white-space:nowrap;')
                                ui.element('div').style('flex:1;height:1px;background:#e8edf2;')

                            with ui.element('div').style('display:flex;flex-wrap:wrap;gap:5px;padding:4px 0;'):
                                for _dk, _dl in [('date_today','今日'), ('date_week','今週'), ('date_month','今月'), ('date_year','今年')]:
                                    with ui.element('div').style(
                                        'display:inline-flex;align-items:center;gap:4px;padding:5px 10px;'
                                        'border-radius:8px;font-size:12px;font-weight:500;cursor:pointer;'
                                        'background:#f0f9ff;color:#0284c7;border:1px solid #bae6fd;'
                                    ) as _dc:
                                        ui.html('<svg style="width:11px;height:11px;flex-shrink:0;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>', sanitize=False)
                                        ui.label(_dl).style('font-size:12px;')
                                    _menu_refs[_dk] = _dc

                            # ── Due-date filter (mode 'w' only) ──
                            if mode_flag == 'w':
                                with ui.element('div').style('display:flex;align-items:center;gap:8px;margin-top:6px;'):
                                    ui.element('div').style('flex:1;height:1px;background:#e8edf2;')
                                    ui.label('完成予定日').style('font-size:10px;color:#94a3b8;font-weight:700;letter-spacing:.05em;white-space:nowrap;')
                                    ui.element('div').style('flex:1;height:1px;background:#e8edf2;')

                                with ui.element('div').style('display:flex;flex-wrap:wrap;gap:5px;padding:4px 0;'):
                                    with ui.element('div').style(
                                        'display:inline-flex;align-items:center;gap:4px;padding:5px 10px;'
                                        'border-radius:8px;font-size:12px;font-weight:500;cursor:pointer;'
                                        'background:#fefce8;color:#ca8a04;border:1px solid #fde68a;'
                                    ) as _dc:
                                        ui.html('<svg style="width:11px;height:11px;flex-shrink:0;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>', sanitize=False)
                                        ui.label('今週完成予定').style('font-size:12px;')
                                    _menu_refs['due_week'] = _dc

                        # "保存済み" sub-divider
                        with ui.element('div').style('display:flex;align-items:center;gap:8px;margin-top:6px;'):
                            ui.element('div').style('flex:1;height:1px;background:#e8edf2;')
                            ui.label('保存済みのユーザーフィルター').style('font-size:10px;color:#94a3b8;font-weight:700;letter-spacing:.05em;white-space:nowrap;')
                            ui.element('div').style('flex:1;height:1px;background:#e8edf2;')

                        # Container for dynamically added saved filter chips
                        with ui.element('div').style('display:flex;flex-direction:column;gap:4px;') as _ffr:
                            with ui.element('div').classes('w-full flex flex-col gap-1') as _fsp:
                                ui.skeleton('QChip').classes('w-full h-8')
                                ui.skeleton('QChip').classes('w-4/5 h-8')
                        _menu_refs['fast_filter_row'] = _ffr
                        _menu_refs['filter_spinner'] = _fsp

                        # Save new filter CTA (dashed border)
                        with ui.element('div').style(
                            'display:flex;align-items:center;justify-content:center;gap:6px;'
                            'width:100%;padding:8px;border-radius:10px;'
                            'border:2px dashed #e2e8f0;font-size:12px;font-weight:500;'
                            'color:#94a3b8;background:none;cursor:pointer;margin-top:6px;'
                        ) as _save_btn:
                            ui.html('<svg style="width:13px;height:13px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M12 4v16m8-8H4"/></svg>', sanitize=False)
                            ui.label('現在のフィルターを保存').style('font-size:12px;')
                        _menu_refs['save_btn'] = _save_btn

                    # Divider
                    ui.element('div').style('height:1px;background:#f0f4f8;margin:4px 12px;')

                    # ── Mark selection section ────────────────────────────
                    with ui.element('div').style(
                        'display:flex;align-items:center;gap:5px;'
                        'font-size:10px;font-weight:700;color:#94a3b8;'
                        'text-transform:uppercase;letter-spacing:.08em;padding:10px 16px 5px;'
                    ):
                        ui.html('<svg style="width:12px;height:12px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 21v-4m0 0V5a2 2 0 012-2h6.5l1 1H21l-3 6 3 6h-8.5l-1-1H5a2 2 0 00-2 2zm9-13.5V9"/></svg>', sanitize=False)
                        ui.label('マーク行を選択')

                    with ui.element('div').style('padding:0 16px 20px;'):
                        _menu_refs['skelist'] = [ui.skeleton('QBtn') for i in range(5)]
                        # Empty-state icon + label
                        with ui.element('div').style('display:flex;flex-direction:column;align-items:center;gap:6px;padding:8px 0;') as mark_empty_state:
                            ui.html('<svg style="width:36px;height:36px;color:#e2e8f0;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M3 21v-4m0 0V5a2 2 0 012-2h6.5l1 1H21l-3 6 3 6h-8.5l-1-1H5a2 2 0 00-2 2zm9-13.5V9"/></svg>', sanitize=False)
                            default_label = ui.label('マーク中の行はありません').style('font-size:11px;color:#94a3b8;')
                        _menu_refs['marks_sub_container'] = ui.row().classes('w-full')
                        _menu_refs['mark_empty_state'] = mark_empty_state
                        mark_empty_state.visible = False
                        skeletons_up(_menu_refs['skelist'], True)

            def add_saved_filter_btn(filter_obj):
                """Add a saved filter chip to fast_filter_row."""
                async def _on_click(fo=filter_obj):
                    await aggrid_instance.apply_filter_model(json.loads(fo.filter_model), valid_columns)
                    column_menu.close()
                async def _on_delete(fo=filter_obj, chip_row=None):
                    result = _DO.delete_filter_memory(USER_INS.uid, fo.filter_id)
                    if result[0] == 's' and chip_row is not None:
                        chip_row.delete()
                    else:
                        ui.notify('削除に失敗しました', type='negative')

                with _menu_refs['fast_filter_row']:
                    with ui.element('div').style('display:flex;align-items:center;gap:4px;') as chip_row:
                        is_global = (filter_obj.db_scope is None)
                        with ui.element('div').style(
                            'display:flex;align-items:center;gap:6px;flex:1;'
                            'padding:6px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                            'background:rgba(86,155,190,.1);color:#3d7fa3;'
                            'border:1px solid rgba(86,155,190,.25);cursor:pointer;'
                        ).on('click', _on_click):
                            ui.html('<svg style="width:12px;height:12px;flex-shrink:0;" fill="currentColor" viewBox="0 0 20 20"><path d="M5 4a2 2 0 012-2h6a2 2 0 012 2v14l-5-2.5L5 18V4z"/></svg>', sanitize=False)
                            ui.label(filter_obj.filter_name).style('font-size:12px;')
                            if is_global:
                                ui.label('全DB').style(
                                    f'font-size:9px;background:rgba(86,155,190,.15);color:{ACCENT_COLOR};'
                                    'padding:1px 5px;border-radius:4px;font-weight:600;margin-left:auto;white-space:nowrap;'
                                )
                        ui.button('✕').style(
                            'width:20px;height:20px;min-width:unset;padding:0;'
                            'background:none;color:#94a3b8;border:none;border-radius:5px;'
                            'font-size:11px;cursor:pointer;'
                        ).props('flat dense').on('click', lambda fo=filter_obj, cr=chip_row: _on_delete(fo, cr))

            async def _lazy_build_menu_shell():
                nonlocal _menu_shell_built, check_row_container, col_ck_sk_list
                if _menu_shell_built:
                    return
                _menu_shell_built = True

                # ── Column checkboxes (heavy UI) ─────────────────────────
                with _menu_refs['col_settings_slot']:
                    check_row_container, col_ck_sk_list = aggrid_instance.init_column_settings(mode_flag)
                _menu_refs['col_settings_spinner'].delete()

                # ── Saved filters (DB call) ──────────────────────────────
                _saved_filters = _DO.get_filter_memories(USER_INS.uid, mode_flag, db_scope)
                for _sf in _saved_filters:
                    add_saved_filter_btn(_sf)
                _menu_refs['filter_spinner'].delete()

                # ── Bind event handlers (first time only) ────────────────
                _menu_refs['preship'].on('click', lambda: fast_filter('preship'))
                _menu_refs['onhold'].on('click', lambda: fast_filter('onhold'))
                _menu_refs['restricted'].on('click', lambda: fast_filter('restricted'))
                if mode_flag in ('i', 'w'):
                    for _k in ('date_today', 'date_week', 'date_month', 'date_year'):
                        _menu_refs[_k].on('click', lambda k=_k: fast_filter(k))
                    if mode_flag == 'w':
                        _menu_refs['due_week'].on('click', lambda: fast_filter('due_week'))
                _menu_refs['save_btn'].on('click', save_current_filter)
                _menu_refs['colapplybt'].on_click(aggrid_instance.update_column)
                _menu_refs['allselectbt'].on_click(lambda: aggrid_instance.column_select('all_on'))
                _menu_refs['allunselectbt'].on_click(lambda: aggrid_instance.column_select('all_off'))

            update_data = ui.button(icon='update', text= f'| 最終更新時刻：{timestamp}').props('color=primary')
            ui.space()
            marked_count = ui.chip(icon='flag', color='lime', text='マークカウント')
            records_count = ui.chip(icon='description', color='orange', text=f'行数: {len(dataframe)}')
            filtered_count = ui.chip(icon='filter_alt', color='blue')
            selected_count = ui.chip(icon='check', color='green')
            marked_count.visible = False
            selected_count.visible = False
            filtered_count.visible = False

        #test = ui.button('test')
        #MARK: AgGrid Initialization
            '''with ui.element('div').classes('w-full h-full flex-grow overflow-hidden'):
                # Create the AgGrid to display the data
                #print(aggrid_column_setting)
                tag_color_scheme = color_tag_rules()
                grid = ui.aggrid.from_pandas(dataframe, options=
                            {
                    'columnDefs' : aggrid_column_setting.copy(),
                    'defaultColDef' : {'flex': 'auto'},
                    'rowSelection': 'multiple',
                    'localeText': japanese_locale_text,
                    'selectAll': 'filtered',
                    'rowClassRules': tag_color_scheme,

                    ':getRowId': '(params) => params.data.lotid',
                                
                }
                ).classes('h-[calc(65vh-2rem)] overscroll-auto')'''
            grid_container, grid = aggrid_instance.aggrid_creator(dataframe)
            if tab_panel is not None:
                tab_panel.aggrid_instance = grid
                tab_panel.aggrid_container = grid_container

            if mode_flag in ('i', 'w'):
                _ctx_row: dict = {}

                def _on_cell_ctx(e):
                    _ctx_row.clear()
                    _ctx_row.update(e.args.get('data', {}))

                async def _on_add_to_schedule():
                    selected = await grid.get_selected_rows() #Fixed Manually
                    rows     = selected if selected else [_ctx_row]
                    lot_ids_sel = [r['lotid'] for r in rows if r.get('lotid')]
                    if not lot_ids_sel:
                        ui.notify('指図番号が取得できません', type='warning')
                        return
                    await _quick_add_to_schedule_dialog(lot_ids_sel, vendor_code, on_save, vendor_name=vendor_name)

                async def _on_reschedule_batch():
                    selected = await grid.get_selected_rows()
                    rows = selected if selected else [_ctx_row]
                    lot_id_flow_product_triples = [
                        (r['lotid'], r.get('flowname', ''), r.get('productcode', ''))
                        for r in rows if r.get('lotid')
                    ]
                    if not lot_id_flow_product_triples:
                        ui.notify('ロットが選択されていません', type='warning')
                        return
                    await _open_reschedule_batch_dialog(lot_id_flow_product_triples)

                grid.on('cellContextMenu', _on_cell_ctx)

            async def update_visbility():
                columnstate = await grid.run_grid_method('getColumnState')
                col_vis_dict = {col['colId'] : not col['hide'] for col in columnstate}
                for col in aggrid_instance.column_list.values():
                    if col.filedid == 'Selected':
                        continue
                    col.visible = col_vis_dict[col.filedid]
                if col_ck_sk_list:
                    skeletons_up(col_ck_sk_list, False)
                if check_row_container is not None:
                    check_row_container.set_visibility(True)
                                

            #test.on_click(printcd)
            async def tag_selected(tagvalue:str, unselect: bool = False):
                selected = await grid.get_selected_rows()
                if not selected:
                    notification('negative', '選択中の行はありません。')
                    return
                selected_ids = tuple(lot['lotid'] for lot in selected)
                tag = tagvalue
                if unselect:
                    tag = '-'
                            
                for id in selected_ids:
                    grid.run_row_method(id, 'setDataValue', 'Selected', tag)

            async def fast_filter(key: str):
                from datetime import datetime, timedelta
                today = datetime.now()
                yesterday         = today - timedelta(days=1)
                last_sunday       = today - timedelta(days=today.weekday() + 1)
                last_month_end    = today.replace(day=1) - timedelta(days=1)
                last_year_end     = today.replace(month=1, day=1) - timedelta(days=1)
                end_of_week       = today + timedelta(days=6 - today.weekday())

                def _d(dt): return dt.strftime('%Y-%m-%d') + ' 00:00:00'

                condition_collection = {
                    'preship' : ('currentprocess',
                                {'filterType': 'text', 'operator': 'OR', 'conditions': [{'filterType': 'text', 'type': 'contains', 'filter': '支給'}, {'filterType': 'text', 'type': 'contains', 'filter': '梱包'}]}),
                    'onhold' : ('onholdreason',
                                {'filterType': 'text', 'type': 'notEqual', 'filter': '-'}),
                    'restricted' : ('restriction',
                                {'filterType': 'text', 'type': 'notEqual', 'filter': '-'}),
                    # ── lastupdatedate period filters ──
                    'date_today':  ('lastupdatedate', {'filterType': 'date', 'type': 'greaterThan', 'dateFrom': _d(yesterday),      'dateTo': None}),
                    'date_week':   ('lastupdatedate', {'filterType': 'date', 'type': 'greaterThan', 'dateFrom': _d(last_sunday),     'dateTo': None}),
                    'date_month':  ('lastupdatedate', {'filterType': 'date', 'type': 'greaterThan', 'dateFrom': _d(last_month_end),  'dateTo': None}),
                    'date_year':   ('lastupdatedate', {'filterType': 'date', 'type': 'greaterThan', 'dateFrom': _d(last_year_end),   'dateTo': None}),
                    # ── completedate filter (mode 'w' which has completedate) ──
                    'due_week':    ('completedate', {'filterType': 'date', 'type': 'inRange',
                                                     'dateFrom': today.strftime('%Y-%m-%d'),
                                                     'dateTo':   end_of_week.strftime('%Y-%m-%d')}),
                }
                condition = condition_collection[key]
                await aggrid_instance.filter_api(condition[0], condition[1])
                column_menu.close()

            async def save_current_filter():
                raw_model = await aggrid_instance.grid.run_grid_method('getFilterModel')
                if not raw_model:
                    ui.notify('フィルターが設定されていません', type='warning')
                    return

                filter_name_holder = {'value': ''}
                global_scope_holder = {'checked': False}  # False = current DB only, True = global

                with ui.dialog().props('backdrop-filter="blur(8px) brightness(40%)"') as save_dialog:
                    with ui.card().style(
                        'border-radius:16px;overflow:hidden;padding:0;min-width:400px;max-width:400px;'
                        'box-shadow:0 20px 60px rgba(0,0,0,.18);'
                    ):
                        # ── Accent header ──────────────────────────────
                        with ui.element('div').style(
                            'display:flex;align-items:center;gap:10px;padding:12px 16px;'
                            f'background:linear-gradient(135deg,{ACCENT_COLOR},#3d7fa3);width:100%;'
                        ):
                            with ui.element('div').style(
                                'width:32px;height:32px;border-radius:50%;'
                                'background:rgba(255,255,255,.2);'
                                'display:flex;align-items:center;justify-content:center;flex-shrink:0;'
                            ):
                                ui.html('<svg style="width:16px;height:16px;color:white;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>', sanitize=False)
                            with ui.element('div'):
                                ui.label('フィルターを保存').style('color:white;font-weight:700;font-size:14px;line-height:1.3;')
                                ui.label('現在のフィルター設定を保存します').style('color:rgba(255,255,255,.6);font-size:11px;')

                        # ── Body ───────────────────────────────────────
                        with ui.element('div').style('padding:20px;display:flex;flex-direction:column;gap:16px;width:100%;box-sizing:border-box;'):

                            # Filter preview box
                            with ui.element('div').style(
                                'background:#f8fafc;border:1px solid #f0f4f8;border-radius:12px;padding:12px;width:100%;box-sizing:border-box;'
                            ):
                                ui.label('保存するフィルター').style(
                                    'font-size:10px;font-weight:700;color:#94a3b8;'
                                    'letter-spacing:.06em;text-transform:uppercase;display:block;margin-bottom:8px;'
                                )
                                _col_label_map = {
                                    'onholdreason': '保留理由', 'currentprocess': '現工程',
                                    'restriction': '制限', 'productcode': '品番',
                                    'lotid': '指図番号', 'flowname': '工程フロー',
                                }
                                _op_map = {
                                    'notEqual': '≠', 'equals': '=', 'contains': '含む',
                                    'notContains': '含まない', 'startsWith': '始まる',
                                    'endsWith': '終わる',
                                    'greaterThan': '>', 'lessThan': '<',
                                    'greaterThanOrEqual': '≥', 'lessThanOrEqual': '≤',
                                    'inRange': '範囲',
                                    'blank': '空白', 'notBlank': '空白でない',
                                }
                                for field, fmodel in (raw_model or {}).items():
                                    col_lbl = _col_label_map.get(field, field)
                                    if 'operator' in fmodel and 'conditions' in fmodel:
                                        op = fmodel['operator']
                                        parts = []
                                        for cond in fmodel.get('conditions', []):
                                            sub_op  = _op_map.get(cond.get('type', ''), cond.get('type', ''))
                                            if cond.get('filterType') == 'date':
                                                raw_d = cond.get('dateFrom', '')
                                                sub_val = raw_d.split(' ')[0] if raw_d else ''
                                            else:
                                                sub_val = cond.get('filter', '')
                                            parts.append(f"{sub_op} {sub_val}".strip())
                                        ftype  = op
                                        op_lbl = op
                                        fval   = f'  {op}  '.join(parts)
                                    else:
                                        ftype = fmodel.get('type', '')
                                        if fmodel.get('filterType') == 'date':
                                            raw_date = fmodel.get('dateFrom', '')
                                            fval = raw_date.split(' ')[0] if raw_date else ''
                                            date_to = fmodel.get('dateTo')
                                            if ftype == 'inRange' and date_to:
                                                fval = f"{fval} 〜 {date_to.split(' ')[0]}"
                                        else:
                                            fval = fmodel.get('filter', '')
                                        op_lbl = _op_map.get(ftype, ftype)
                                    with ui.element('div').style(
                                        'display:flex;align-items:center;gap:6px;padding:3px 0;'
                                    ):
                                        ui.element('span').style(f'width:6px;height:6px;border-radius:50%;background:{ACCENT_COLOR};flex-shrink:0;display:inline-block;')
                                        ui.label(col_lbl).style('font-size:11px;font-weight:500;color:#64748b;')
                                        ui.label(op_lbl).style('font-size:11px;color:#94a3b8;')
                                        ui.label(str(fval)).style(
                                            'font-family:monospace;font-size:11px;background:white;'
                                            'padding:1px 6px;border-radius:5px;border:1px solid #e8edf2;'
                                        )

                            # Filter name input
                            with ui.element('div').style('width:100%;box-sizing:border-box;'):
                                ui.label('フィルター名').style(
                                    'display:block;font-size:11px;font-weight:700;color:#64748b;'
                                    'letter-spacing:.03em;margin-bottom:6px;'
                                )
                                filter_name_input = ui.input(placeholder='例：保留中 スパッタ工程').style(
                                    'width:100%;'
                                ).props('outlined dense')
                                filter_name_input.bind_value(filter_name_holder, 'value')

                            # Scope selector (only for mode 'w')
                            if mode_flag == 'w':
                                with ui.element('div').style('width:100%;box-sizing:border-box;'):
                                    ui.label('適用範囲').style(
                                        'display:block;font-size:11px;font-weight:700;color:#64748b;'
                                        'letter-spacing:.03em;margin-bottom:6px;'
                                    )
                                    scope_current_el = None
                                    scope_global_el = None

                                    def _set_scope(is_global: bool):
                                        global_scope_holder['checked'] = is_global
                                        if is_global:
                                            scope_current_el.style(
                                                'flex:1;padding:9px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                                                'color:#64748b;cursor:pointer;text-align:center;'
                                                'border:1.5px solid #e2e8f0;user-select:none;'
                                            )
                                            scope_global_el.style(
                                                'flex:1;padding:9px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                                                'color:#3d7fa3;cursor:pointer;text-align:center;'
                                                f'background:rgba(86,155,190,.12);border:1.5px solid {ACCENT_COLOR};'
                                            )
                                        else:
                                            scope_current_el.style(
                                                'flex:1;padding:9px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                                                'color:#3d7fa3;cursor:pointer;text-align:center;'
                                                f'background:rgba(86,155,190,.12);border:1.5px solid {ACCENT_COLOR};'
                                            )
                                            scope_global_el.style(
                                                'flex:1;padding:9px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                                                'color:#64748b;cursor:pointer;text-align:center;'
                                                'border:1.5px solid #e2e8f0;user-select:none;'
                                            )

                                    with ui.element('div').style('display:flex;gap:8px;width:100%;'):
                                        with ui.element('div').style(
                                            'flex:1;padding:9px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                                            'color:#3d7fa3;cursor:pointer;text-align:center;'
                                            f'background:rgba(86,155,190,.12);border:1.5px solid {ACCENT_COLOR};'
                                        ).on('click', lambda: _set_scope(False)) as scope_current_el:
                                            ui.label('このDB品番のみ').style('font-size:12px;font-weight:500;')
                                            ui.label(db_scope or '').style(f'font-size:10px;color:{ACCENT_COLOR};')
                                        with ui.element('div').style(
                                            'flex:1;padding:9px 10px;border-radius:10px;font-size:12px;font-weight:500;'
                                            'color:#64748b;cursor:pointer;text-align:center;'
                                            'border:1.5px solid #e2e8f0;'
                                        ).on('click', lambda: _set_scope(True)) as scope_global_el:
                                            ui.label('全DB品番').style('font-size:12px;font-weight:500;')
                                            ui.label('グローバル').style('font-size:10px;color:#94a3b8;')

                        # ── Footer ────────────────────────────────────
                        with ui.element('div').style(
                            'display:flex;align-items:center;justify-content:space-between;gap:12px;'
                            'padding:14px 20px;border-top:1px solid #f1f5f9;background:rgba(248,250,252,.7);'
                            'width:100%;box-sizing:border-box;'
                        ):
                            ui.button('キャンセル', on_click=save_dialog.close).style(
                                'padding:9px 18px;background:transparent;color:#64748b;'
                                'border-radius:9px;font-size:13px;font-weight:500;border:1.5px solid #e2e8f0;'
                            ).props('flat')

                            async def do_save():
                                if not filter_name_holder['value']:
                                    ui.notify('フィルター名を入力してください', type='warning')
                                    return
                                scope = None
                                if mode_flag == 'w':
                                    scope = None if global_scope_holder['checked'] else db_scope
                                result = _DO.save_filter_memory(
                                    USER_INS.uid,
                                    filter_name_holder['value'],
                                    mode_flag,
                                    scope,
                                    json.dumps(raw_model)
                                )
                                save_dialog.close()
                                if result[0] == 's':
                                    ui.notify('フィルターを保存しました', type='positive')
                                    try:
                                        new_fm = _DO.FilterMemory()
                                        new_fm.namecode = USER_INS.uid
                                        new_fm.filter_id = result[1]
                                        new_fm.filter_name = filter_name_holder['value']
                                        new_fm.query_mode = mode_flag
                                        new_fm.db_scope = scope
                                        new_fm.filter_model = json.dumps(raw_model)
                                        add_saved_filter_btn(new_fm)
                                    except Exception:
                                        pass
                                else:
                                    ui.notify(f'保存に失敗しました: {result[1]}', type='negative')

                            with ui.button(on_click=do_save).style(
                                f'padding:9px 18px;background:{ACCENT_COLOR};color:white;'
                                'border-radius:9px;font-size:13px;font-weight:600;'
                            ):
                                ui.html('<svg style="width:14px;height:14px;margin-right:4px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>', sanitize=False)
                                ui.label('保存する')

                save_dialog.open()

            async def filter_by_context():
                if CLICK_EVENT_CACHE is None:
                    notification('negative', 'テーブル範囲外のエリアが選択されました。')
                    return
                current_cell: str = CLICK_EVENT_CACHE.args['value']
                current_col = CLICK_EVENT_CACHE.args['colId']
                filter_condition = None
                if 'date' in current_col:
                    filter_condition = {'filterType': 'date', 'type': 'equals', 'dateFrom': f'{current_cell.split(' ')[0]} 00:00:00', 'dateTo': None}
                else:
                    filter_condition = {'filterType': 'text', 'type': 'equals', 'filter': current_cell}
                grid.run_grid_method('setColumnFilterModel',
                                                current_col,  
                                                filter_condition)
                grid.run_grid_method('onFilterChanged')

            async def read_mark_color(gen: int = 0):
                data: List[Dict[str, str | int]]  = await grid.get_client_data()
                if gen and gen != _menu_gen:   # menu was closed while we awaited — bail out
                    return
                #data: = grid.options['rowData']
                count_marks = Counter([mark['Selected'] for mark in data])
                count_marks.pop('-', None)
                column_menu.data = count_marks

                with _menu_refs['marks_sub_container']:
                    color_list: Counter = column_menu.data
                    skeletons_up(_menu_refs['skelist'], False)

                    if color_list:
                        _menu_refs['mark_empty_state'].visible = False
                        with ui.row(align_items='start').classes('w-full'):
                            for color, count in color_list.items():
                                ui.checkbox(f'{color} = {count}')
                        with ui.row(align_items='start').classes('w-full'):
                            ui.button('選択', on_click=_selectallrows_marked)
                            #ui.button('te', on_click=get_selected_marks)
                    else:
                        _menu_refs['mark_empty_state'].visible = True

            async def menu_pre_process():
                nonlocal _menu_gen
                _menu_gen += 1
                gen = _menu_gen
                await _lazy_build_menu_shell()
                if gen != _menu_gen:
                    return
                await update_visbility()
                if gen != _menu_gen:
                    return
                await read_mark_color(gen)


            async def clear_menu():
                nonlocal _menu_gen
                _menu_gen += 1          # invalidate any in-flight menu_pre_process
                if not _menu_shell_built:
                    return
                check_row_container.set_visibility(False)
                skeletons_up(col_ck_sk_list, True)
                _menu_refs['marks_sub_container'].clear()
                skeletons_up(_menu_refs['skelist'], True)
                _menu_refs['mark_empty_state'].visible = False

            def get_selected_marks() -> List[str]:
                firstrow = _menu_refs['marks_sub_container'].default_slot.children[0]
                markedlist = []
                for i in firstrow.default_slot.children:
                    if i.props['model-value']:
                        markedlist.append(i._text.split()[0])
                return markedlist
                                
            
            with ui.context_menu() as cmenu:
                copymenu = ui.menu_item('選択行をコピー')
                selcolmenu = ui.menu_item('列を選択/解除')
                deselectallmenu = ui.menu_item('列選択全解除')
                ui.separator()
                detailmenu = ui.menu_item('詳細', on_click=lambda detailmenu_caller = create_process_dialog: context_wrapper(detailmenu_caller))
                serialmenu = ui.menu_item('シリアル')
                ui.separator()
                with ui.menu_item('マーク', auto_close=False) as mark_menu:
                    with ui.item_section().props('side'):
                        ui.icon('keyboard_arrow_right')
                    with ui.menu().props('anchor="top end" self="top start" auto-close'):
                        for tag in TAG_COLOR:
                            tag_menu_item = ui.menu_item(tag, on_click=lambda t=tag: grid.run_row_method(get_click_cache().args['rowId'], 'setDataValue', 'Selected', t))
                            with tag_menu_item:
                                ui.label('　　⬤').classes(f'text-[{TAG_COLOR[tag]}]')
                        ui.menu_item('取り消し', on_click=lambda t=tag: grid.run_row_method(get_click_cache().args['rowId'], 'setDataValue', 'Selected', '-'))
                with ui.menu_item('選択中の行をマーク', auto_close=False):
                    with ui.item_section().props('side'):
                        ui.icon('keyboard_arrow_right')
                    with ui.menu().props('anchor="top end" self="top start" auto-close'):
                        for tag in TAG_COLOR:
                            batch_tag_menu_item = ui.menu_item(tag, on_click=lambda t=tag: tag_selected(t))
                            with batch_tag_menu_item:
                                ui.label('　　⬤').classes(f'text-[{TAG_COLOR[tag]}]')
                        ui.menu_item('取り消し', on_click=lambda t=tag: tag_selected(t, True))

                ui.separator()
                filter_by_con = ui.menu_item('選択中セルの値でフィルター', on_click=filter_by_context)
                ui.separator()
                excel_write_menu = ui.menu_item('Excelへの書き出し')
                if mode_flag in ('i', 'w'):
                    ui.separator()
                    ui.menu_item('支給日程表に追加', on_click=_on_add_to_schedule)
                    ui.menu_item('リスケ依頼生成（バッチ）', on_click=_on_reschedule_batch)

            column_menu.on('show', menu_pre_process)
            column_menu.on('hide', clear_menu)
            cmenu.on('before-show', clear_event_cache)
            grid.on('cellContextMenu', click_cache)

            grid.on('cellDoubleClicked', context_adapted_query)
            selected_count.on_click(lambda: grid.run_grid_method('deselectAll'))
            filtered_count.on_click(lambda: grid.run_grid_method('setFilterModel', 'null'))
            #grid.on('cellContextMenu', lambda event: right_click_cache(event))

   #         #AGGRID_SELECTION[grid.id] = aggrid_instance.column_list

            async def sel_rows():
                a = await grid.get_selected_rows()
                print(a)

            #test.on_click(sel_rows)

        # Add export buttons or other functionality
        with ui.row().classes('w-full justify-end gap-2 mt-4'):
            ui.button('CSVに出力', icon='file_download').props('color=primary').on('click', lambda: export_csv(dataframe, search_mode["value"]))
            #ui.button('Print', icon='print').on('click', lambda: print_grid(tab_id))

    #grid.update()

        #test_data = {'lotid': [1, 2], 'productcode': [3, 4], 'flowname': [3, 4], 'currentprocess': [3, 4], 'inprogress': [3, 4], 
        #             'lastupdatedate': [3, 4], 'SEQ': [3, 4], 'onhold': [3, 4], 'onholdreason': [3, 4], 'completedate': [3, 4],
        #             'StartStatus': [3, 4], 'restriction': [3, 4]}
        
        #test_df = pd.DataFrame(data= test_data)
    else:
        #MARK: Relation Tree initialization
        tree_data = M.tree_parser(dataframe)
        with ui.row().classes('w-full items-center flex-nowrap'):
            tree_toggle_style = 'color=secondary text-color=primary toggle-text-color=secondary toggle-color=primary'
            ui.label('出力内容： ')
            include_toggle = ui.toggle({'all': '指図 : 品番', 'lot' : '指図のみ'}, value= 'all').props(tree_toggle_style)
            #ui.separator().props('vertical self-center')
            ui.label('・　出力形式： ')
            format_toggle = ui.toggle({'horizontal' : '横', 'vertical' : '縦', 'tree' : '木構造'}, value= 'horizontal').props(tree_toggle_style)
            #ui.separator().props('vertical')
            ui.label(f'・　関係図レコード数：{len(tree_data)}')
            ui.space()
            l_select_condition = ui.chip(icon='rule', color='orange-4', text='条件付き選択')
            l_select_all = ui.chip(icon='checklist', color='orange', text='リスト全選択')
            l_unselect_all = ui.chip(icon='cancel', color='grey', text='リスト全取消')
            l_copy = ui.chip(icon='content_copy', color='green', text='リストコピー')
                    
        with l_select_condition:
            with CustomMenu() as condition_menu:
                with ui.card() as condition_container:
                    ui.label('条件付き選択')
                    db = ui.input('DB図番')
                    ui.label('OR')
                    search_type = ui.input('タイプ')
                    condition_c_search_button = ui.label('検索').classes('cursor-pointer bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-200 p-2 hover:shadow-2xl hover:shadow-slate-300/50 transition-all duration-500 text-sm')

        async def get_tick_nodes(tree: ui.tree, list_: bool = False) -> str:
            include = include_toggle.value
            mode = format_toggle.value
            ticked = await tree.run_method('getTickedNodes')
            if not ticked:
                if not list_:
                    notification('negative', '選択中の階層はありません。')
                return
            else:
                holder_dict = {}
                for node in ticked:
                    node_str = node['id']
                    if include == 'all':
                        node_str += f': {node['description']}'
                    holder_dict[node['level']] = node_str
            #holder_list = (holder_dict[node_key] for node_key in sorted(holder_dict.keys()))
            holder_list = (v for _, v in sorted(holder_dict.items()))

            if mode == 'horizontal':
                return ' <- '.join(holder_list)
            if mode == 'vertical':
                return '\n\t  |  \n'.join(holder_list)
            if mode == 'tree':
                holder = '├ '
                seperator = '│'
                space = '\t'
                for n in holder_list:
                    if holder == '├ ':
                        holder += n + '\n'
                    else:
                        holder += f'{seperator}{space}└ {n}\n'
                        space += '\t'
                return holder
                    
        async def copy_ticked_nodes(tree: ui.tree):
            result = await get_tick_nodes(tree)
            if result:
                if M.singlecopy(result):
                    notification('positive', '選択した部分をコピーしました。')

        async def copy_tree_list(tree_list: List[ui.tree]):
            holder_str =''
            success_count = 0
            for t in tree_list:
                tree_str = await get_tick_nodes(t, True)
                if tree_str:
                    success_count += 1
                    if holder_str != '':
                        holder_str += '\n\n'
                    holder_str += tree_str
            if holder_str != '' and M.singlecopy(holder_str):
                notification('positive', f'関係図 {success_count}件をコピーしました。')
                    
                    

        tree_list = []
                    
        for t in tree_data:
            with ui.card().classes('w-full rounded-lg border-2 border-indigo-200 border-b-indigo-500'):
                with ui.row().classes('w-full p-px item-center'):
                    with ui.chip(color='white'):
                        ui.label('検索対象指図').classes('underline decoration-sky-500 decoration-2 text-lg')
                    ui.space()
                    checkchip = ui.chip(icon='checklist', color='orange', text='全選択')
                    cancelchip = ui.chip(icon='cancel', color='grey', text='全取消')
                    copy_chip_t = ui.chip(icon='content_copy', color='green', text='コピー')
                ui.separator()
                tree = ui.tree([t], label_key='id', tick_strategy='strict').classes('p-px')
                tree_list.append(tree)
                checkchip.on_click(tree.tick)
                cancelchip.on_click(tree.untick)
                copy_chip_t.on_click(lambda t= tree: copy_ticked_nodes(t))
                keys_for_highlight = refresh_clean_key.key

                tree.add_slot('default-header', f'''
                    <span :props="props"
                            :class="{{
                            'highlighted-node underline decoration-sky-500 decoration-2': [{",".join(tuple("'" + k + "'" for k in keys_for_highlight))}].includes(props.node.id)
                            }}"         
                              >指図: <strong>{{{{ props.node.id }}}}</strong></span>
                ''')
                tree.add_slot('default-body', '''
                    <span :props="props">DB図番: "{{ props.node.description }}", 階層: {{ props.node.level}}, </span>
                    <span :props="props">中間品種類: "{{ props.node.type }}" </span>
                ''')

                tree.run_method('expandAll')
        #tree.expand()
        l_select_all.on_click(lambda tl= tree_list: tick_all_tree(tl, True))
        l_unselect_all.on_click(lambda tl= tree_list: tick_all_tree(tl, False))
        l_copy.on_click(lambda tl= tree_list: copy_tree_list(tl))
        l_select_condition.on_click(condition_menu.open)
        condition_c_search_button.on('click', 
        lambda d= db, s=search_type: M.search_tree(tree_list, d, s))

    #MARK: AgGrid Helper Functions
    async def update_select_count(e):
        global END_OF_BATCH
        #print(e.args, e, END_OF_BATCH)
        if e.args['source'] == 'api' and not END_OF_BATCH:
            return
        else:
            END_OF_BATCH = None
        selected_rows = await grid.get_selected_rows()
        selected = len(selected_rows)
        if selected < 1:
            selected_count.visible = False
            return
        selected_count.text = f'選択数: {selected}'
        selected_count.visible = True
        

    async def update_row_data():
        spinner_open()
        df_cluster = await uni_query(refresh_clean_key)
        client_data = await grid.get_client_data()
        og_data = {row['lotid']: row['Selected'] for row in client_data}
            
        #df_cluster['Selected'] = df_cluster['lotid'].map(og_data)
        def is_special_dtype(dtype):
            return (pd.api.types.is_datetime64_any_dtype(dtype) or
                    pd.api.types.is_timedelta64_dtype(dtype) or
                    pd.api.types.is_complex_dtype(dtype) or
                    isinstance(dtype, pd.PeriodDtype))
        df = tuple(df_cluster.values())[0]
        df['Selected'] = df['lotid'].map(og_data)
        special_cols = df.columns[df.dtypes.apply(is_special_dtype)]
        if not special_cols.empty:
            df = df.copy()
            df[special_cols] = df[special_cols].astype(str)

        if isinstance(df.columns, pd.MultiIndex):
            raise ValueError('MultiIndex columns are not supported. '
                            'You can convert them to strings using something like '
                            '`df.columns = ["_".join(col) for col in df.columns.values]`.')



        coverted_df = df.to_dict('records')
        #grid.options['rowData'] = coverted_df

        timestamp = datetime.now().strftime('%H:%M:%S')
        update_data.text = f'| 最終更新時刻：{timestamp}'
            
        aggrid_data_update(grid, coverted_df)

        spinner_close()

        notification('positive', 'データを更新しました。')


    #test.on_click(lambda: update_row_data(test_df))

    async def multiple_line_serial():
        key = await grid.get_selected_rows()
        if not key:
            key: dict = get_click_cache().args['data']
            #key = (c_cache.args['data'] if c_cache is not None else None,)
            #print(type(key), key)
            if key.get('lotid', None) is None:
                notification('negative', '選択中の行はありません。')
                raise E.NoKeyExc
            key = [key]
        key_list = tuple(r['lotid'] for r in key)
        result_frame, _ = await query_data_stdin_no_validation('s', key_list)
        await create_process_dialog(serial_= result_frame)
        


    #def left_click_event(event):
    #    click_cache(event)
    #    if KEY_CACHE.alt:
    #        highlight_column()

    async def get_filter_count():
        filtered_model = await grid.run_grid_method('getFilterModel')
        if filtered_model:
            selected = await grid.get_client_data(method='filtered_unsorted')
            filtered_count.text = f'フィルター内:{len(selected)}'
            filtered_count.visible = True
        else:
            filtered_count.visible = False
            return

    #async def _selectallrows():
    #    global END_OF_BATCH
    #    filtered_model = await grid.run_grid_method('getFilterModel')
    #    if not filtered_model:
    #        await grid.run_grid_method('(grid) => grid.selectAll()')
    #    else:
    #        selected = await grid.get_client_data(method='filtered_unsorted')
    #        selected_id = [row['lotid'] for row in selected]
    #        END_OF_BATCH = False
    #        for id in selected_id:
    #            grid.run_row_method(id, 'setSelected', True)
    #        END_OF_BATCH = True

    async def _selectallrows_marked():
        global END_OF_BATCH
        marks = get_selected_marks()
        filtered_model = await grid.run_grid_method('getFilterModel')
        selected = await grid.get_client_data(method='filtered_unsorted')
        selected_id = [row['lotid'] for row in selected if row['Selected'] in marks]
        END_OF_BATCH = False
        for id in selected_id:
            grid.run_row_method(id, 'setSelected', True)
        END_OF_BATCH = True
        notification('positive', f'マーク{','.join(marks)}の行を選択しました。')
        column_menu.close()       


    #def _deselectallrows():
    #    grid.run_grid_method('deselectAll', 'filtered')

    #def _deselect():
    #    for col in grid.options['columnDefs']:
    #        col['headerClass'] = 'None'
    #    deselect_all(grid.id)
    #    aggrid_update(grid, grid.options)
        #grid.update()

    #async def awaitrowdata():
    #    rows = await grid.get_selected_rows()
    #    selected = get_selected_cols(rows, grid.id)
    #    if rows:
    #        if M.copytoclipboard(selected):
    #            notification('positive', '行をコピーしました。')

    #async def keyboard_copy():
    #    if KEY_CACHE.key == 'c' and KEY_CACHE.ctrl:
    #        await awaitrowdata()
    #    if KEY_CACHE.key == 'c' and KEY_CACHE.alt:
    #        copy_cell()
    #    if KEY_CACHE.key == 'a' and KEY_CACHE.ctrl:
    #        await _selectallrows()
        #if KEY_CACHE.key == 's' and KEY_CACHE.ctrl:
        #    await multiple_line_serial()
        #need to figure a way to suppress the browser event
        
    def sel_column():
            
        global CLICK_EVENT_CACHE
        e = CLICK_EVENT_CACHE
        if not e:
            return
        col_name = e.args['colId']
        #print(col_name)
        se#lection = AGGRID_SELECTION[grid.id]
        ind = indexed_column_header[col_name][0]
        attribute_target = grid.options['columnDefs'][ind]
        if 'hide' not in attribute_target or not attribute_target['hide']:
            attribute_target['hide'] = True
            selection[col_name].flick(False)
            #print(f'colname => {col_name}\n\n')
            #print(selection)
        aggrid_update(grid, grid.options)
        #grid.update()

    #def highlight_column():
            
    #    global CLICK_EVENT_CACHE
    #    e = CLICK_EVENT_CACHE
    #    if not e:
    #        return
    #    col_name = e.args['colId']
        #print(col_name)
    #    se#lection = AGGRID_SELECTION[grid.id]
    #    ind = indexed_column_header[col_name][0]
    #    attribute_target = grid.options['columnDefs'][ind]
        #print(grid.options)
    #    if 'headerClass' not in attribute_target or attribute_target['headerClass'] != 'selected-column':
    #        attribute_target['headerClass'] = 'selected-column'
    #        selection[col_name].highlight_flip(True)
    #    elif attribute_target['headerClass'] == 'selected-column':
    #        attribute_target['headerClass'] = 'None'
    #        selection[col_name].highlight_flip(False)
        #grid.update()
    #    aggrid_update(grid, grid.options)
        #print('after')
        #print(grid.options)

    #def update_column():
        #for col in column_list.keys():
        #    grid.options['columnDefs'][column_list[col].colid]['hide'] = not column_list[col].visible
    #    for col in column_list.values():
    #        grid.options['columnDefs'][col.colid]['hide'] = not col.visible
    #    aggrid_update(grid, grid.options)
        #grid.update()

    #def copy_cell():
    #    global CLICK_EVENT_CACHE
    #    if CLICK_EVENT_CACHE:
    #        if M.singlecopy(CLICK_EVENT_CACHE.args['value']):
    #            CLICK_EVENT_CACHE = None
    #            notification('positive', 'セルをコピーしました。')

    #MARK: Relation Tree screen Binding
    if mode_flag != 't':
        update_data.on_click(update_row_data)
        grid.on('selectionChanged', update_select_count)
        grid.on('filterChanged', get_filter_count)
        serialmenu.on_click(multiple_line_serial)
        #grid.on('cellKeyDown', keyboard_copy)
        #grid.on('cellClicked', left_click_event)
        aggrid_instance.cellkey_func_bind()
        copymenu.on_click(aggrid_instance.awaitrowdata)

        async def open_excel_write():
            rows = await grid.get_selected_rows()
            data = aggrid_instance.get_selected_cols(rows)
            if not rows:
                notification('warning', '行を選択してから実行してください。')
                return
            await excel_write_dialog(data)

        excel_write_menu.on_click(open_excel_write)
        selcolmenu.on_click(aggrid_instance.highlight_column)
        deselectallmenu.on_click(aggrid_instance._deselect)
    #else:
        #copy_chip_t.on_click(get_tick_nodes)
        #pass
        
    #selectiontest.on_click(test_get_first)



async def newtab(dataframe: pd.DataFrame | List[Tuple[str]], name: str, refresh_clean_key: M.query_, mode_flag_:str = None, timestmp: str =None) -> Tuple[ui.tab, ui.tab_panel]:
    #dataframe['#'] = [0] * len(dataframe)
    #print(dataframe)
    if timestmp is None:
        timestamp = datetime.now().strftime('%H:%M:%S')
    else:
        timestamp = timestmp
    if mode_flag_ is None:
        mode_flag = search_mode['value']
    else:
        mode_flag = mode_flag_

    # db_scope: None for type 'i' (always global); DB code string for type 'w'
    db_scope = None
    if mode_flag == 'w' and refresh_clean_key.key:
        db_scope = refresh_clean_key.key[0] if isinstance(refresh_clean_key.key, list) else refresh_clean_key.key

    #if mode_flag != 't':
    #    aggrid_column_setting = generate_col_settings(mode_flag)
    #    indexed_column_header = {headername['field'] : (index, headername['headerName']) for index, headername in enumerate(
    #        col for col in aggrid_column_setting 
    #    )}
        '''indexed_column_header_name = {headername : index for index, headername in enumerate(
            col['headerName'] for col in aggrid_column_setting 
        )}'''
        #print(indexed_column_header)

    name = '指図進捗' if 'i' in name else '指図関係' if 't' in name else name

    #print(dataframe)

    for tab in tabs.default_slot:
        if name in str(tab):
            #tabs.remove(tab)
            try:
                tab_delete(tab, tab.id)
            except KeyError as KE:
                notification('negative', f'タブ削除に失敗しました。KEYが見つかりませんでした:\n{KE}', multiline= True)
                #print(KE, KE.with_traceback, KE.__dir__)
                spinner_close()
                return
            except Exception as e:
                notification('negative', f'タブ削除に失敗しました。予定外のエラーが発生しました:\n{e}', multiline= True)
                #print(e, e.with_traceback, e.__dir__)
                spinner_close()
                return
            #tab.delete()
            #print(f'removed {name}')
    
    newtabname = name

    def _build_inventory_viewer(db_code: str) -> None:
        """在庫数量 dropdown for 仕掛 tabs.
        Stage 1 (menu show): loads inventory_combined with skeleton, populates rows 1+2.
        Stage 2 (部品 show): lazy-loads inventory_components into expansion body.
        """
        import asyncio
        _inv_loaded   = False
        _compo_loaded = False

        _ROW  = 'display:flex;justify-content:space-between;align-items:center;padding:10px 16px;gap:24px;cursor:default;'
        _CODE = 'font-family:monospace;font-size:13px;color:#555;white-space:nowrap;'
        _QTY  = 'flex-shrink:0;'

        with ui.button('在庫数量', icon='inventory_2').props(
                'flat outline color=primary no-caps'):
            with ui.menu().style('width:520px').props('anchor="bottom right" self="top right"') as inv_menu:

                # ── Row 1: current product ──────────────────────────────
                with ui.element('div').style(_ROW):
                    ui.label(db_code).style(_CODE)
                    inv_qty_slot = ui.element('div').style(_QTY + 'display:flex;flex-direction:column;gap:2px;align-items:flex-end;')
                    with inv_qty_slot:
                        ui.skeleton(type='text').props('width="40px"')

                ui.separator()

                # ── Row 2: WIP previous stage ───────────────────────────
                with ui.element('div').style(_ROW):
                    wip_name_slot = ui.element('div').style('white-space:nowrap;')
                    with wip_name_slot:
                        ui.skeleton(type='text').props('width="80px"')
                    wip_qty_slot = ui.element('div').style(_QTY + 'display:flex;flex-direction:column;gap:2px;align-items:flex-end;')
                    with wip_qty_slot:
                        ui.skeleton(type='text').props('width="40px"')

                ui.separator()

                # ── Row 3: components (lazy expansion) ─────────────────
                with ui.expansion('部品', icon='category').classes(
                        'w-full text-sm').props('content-style="padding:0"') as compo_exp:
                    compo_list_container = ui.element('div').style('width:100%')

        _PLANT_LABEL = {'5501': '伊勢原', '5401': '宮田'}

        async def _load_inv_and_wip() -> None:
            nonlocal _inv_loaded
            if _inv_loaded:
                return
            _inv_loaded = True

            rows = await M.fast_query('inventory_combined', [db_code, db_code]) or []
            current_rows = [r for r in rows if r[0] == 'CURRENT']
            wip_rows     = [r for r in rows if r[0] == 'WIP']

            inv_qty_slot.clear()
            with inv_qty_slot:
                if current_rows:
                    for r in current_rows:
                        plant = _PLANT_LABEL.get(str(r[3]), str(r[3])) if r[3] else ''
                        label_text = f'{plant}: {r[2]}' if plant else str(r[2])
                        ui.label(label_text).style('font-weight:700;color:#1565c0;font-size:14px;')
                else:
                    ui.label('—').style('font-weight:700;color:#1565c0;font-size:14px;')

            wip_qty_slot.clear()
            wip_name_slot.clear()
            if wip_rows:
                with wip_name_slot:
                    for r in wip_rows:
                        ui.label(str(r[1])).style(_CODE)
                with wip_qty_slot:
                    for r in wip_rows:
                        plant = _PLANT_LABEL.get(str(r[3]), str(r[3])) if r[3] else ''
                        label_text = f'{plant}: {r[2]}' if plant else str(r[2])
                        ui.label(label_text).style('font-weight:700;color:#2e7d32;font-size:14px;')
            else:
                with wip_name_slot:
                    ui.label('-').style(_CODE)
                with wip_qty_slot:
                    ui.label('0').style('font-weight:700;color:#2e7d32;font-size:14px;')

        async def _load_components() -> None:
            nonlocal _compo_loaded
            if _compo_loaded:
                return
            _compo_loaded = True

            with compo_list_container:
                skel = ui.skeleton(type='text').props('width="100%"')
            rows = await M.fast_query('inventory_components', [db_code]) or []
            skel.delete()
            with compo_list_container:
                if not rows:
                    ui.label('部品なし').style('font-size:11px;color:#bbb;padding:8px 16px;display:block;')
                    return
                for row in rows:
                    with ui.element('div').style(
                        'display:flex;justify-content:space-between;align-items:center;'
                        'padding:6px 16px;gap:16px;border-top:1px solid #f0f0f0;width:100%;'
                    ):
                        display_name = str(row[1])
                        plant = _PLANT_LABEL.get(str(row[3]), str(row[3])) if row[3] else ''
                        qty_text = f'{plant}: {row[2]}' if plant else str(row[2])
                        ui.label(display_name).style('font-family:monospace;font-size:11px;color:#555;white-space:nowrap;')
                        ui.label(qty_text).style('font-size:11px;font-weight:700;color:#e65100;flex-shrink:0;')

        inv_menu.on('show', lambda: asyncio.ensure_future(_load_inv_and_wip()))
        compo_exp.on('show', lambda: asyncio.ensure_future(_load_components()))

    with tabs:
        color = ''
        if '仕掛' in newtabname:
            color = 'green'
        if '進捗' in newtabname:
            color = 'sky'
        if '関係' in newtabname:
            color = 'amber'
        new_tab = ui.tab(newtabname).classes(f'rounded-t-lg bg-{color}-200').props('closable')
        #print(f'newtab id ={new_tab.id}')
        with new_tab:
            with ui.context_menu():
                ui.menu_item('削除', on_click= lambda: tab_delete(new_tab, new_tab.id))

    #column_list = {}
    with tab_panels:
        with CustomTabPanel(new_tab).classes('w-full h-full') as new_tab_panel:
            #print(f'new_tab_panel id ={new_tab_panel.id}'
            with ui.row().classes('items-center gap-3 mb-2 w-full'):
                ui.label(f'{newtabname} 検索結果').classes('text-lg font-bold')
                if '仕掛' in newtabname:
                    ui.space()
                    _build_inventory_viewer(newtabname.split('_', 1)[1])
            with ui.card().classes(f'w-full h-full flex flex-col bg-{color}-50') as container_card:
                _render_result_panel(
                    container_card, dataframe, mode_flag, newtabname,
                    refresh_clean_key, db_scope=db_scope, timestamp=timestamp,
                    tab_panel=new_tab_panel,
                )
    current_tabs[new_tab.id] = ((dataframe, name, refresh_clean_key, mode_flag, timestamp), new_tab, new_tab_panel)
    search_prompt['prompt'] = False
    return new_tab, new_tab_panel




# Function to perform the search and update the UI
async def perform_search(textinput: ui.textarea, dialog: ui.dialog):
    #import asyncio
    #import time
    # Get the data from the query
    #timer_start = time.perf_counter()

    df, reuse_search_parameters = await query_data(textinput, dialog)
    #end_of_search = time.perf_counter() - timer_start

    #print(f'Search completed in {end_of_search:.2f} seconds')
    #print(df)
    
    mode_str = {
        'i' : '指図',
        'w' : '仕掛',
        't' : '関係'
    }
    #count_ = 0

    mode_flag = search_mode['value']

    #start_of_tab_creation = time.perf_counter()
    for key, val in df.items():
        #print(f'iteration = {count_}')
        data_frame = val
        if mode_flag == 'w':
            new_search_parameters = reuse_search_parameters.copy([key])
        else:
            new_search_parameters = reuse_search_parameters.copy()

        ui.context.client.check_existence()
        new_tab, _ = await newtab(data_frame, f'{mode_str[mode_flag]}_{key}', new_search_parameters)
        #print(f'Tab for {key} created in {time.perf_counter() - start_of_tab_creation:.2f} seconds')

        #count_ += 1
    
    # Close the dialog after search
    dialog.close()
    spinner_close()
    
    # Activate the new tab
    try:
        tabs.set_value(new_tab)
    except Exception as e:
        ui.notify(f'Error activating tab: {str(e)}', type='negative')
        #print(f"Tab activation error: {str(e)}")

# Function to export data to CSV
def export_csv(df: pd.DataFrame, mode: str):
    if df is not None:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = os.path.join(desktop, f'{mode}_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        os.startfile(filename)
        ui.notify('データをデスクトップにCSVファイルで出力しました。', type='positive')
    else:
        ui.notify('CSVに出力できる有効なデータはありません。', type='negative')

# Function to print grid (placeholder)
def print_grid(tab_id):
    ui.notify('Print functionality would be implemented here', type='info')

def create_custom_table(data: pd.DataFrame | List[Tuple[str]]) -> Tuple[ui.column, CustomTable]:
    """
    Create a custom table from a DataFrame or list of tuples.
    """
    new_table: CustomTable = None
    def update_table(switch: ui.switch):
        set_to_value = 'None'
        if not switch.value:
            set_to_value = 'selected-column-table'
        col_obj = new_table.indexed_column_name_dict[switch.text]
        select_column(col_obj, set_to_value)
        #for col in new_table.columns:
        #    if col['name'] == switch.text and not switch.value:
        #        col['classes'] = 'selected-column-table'
        #        col['headerClasses'] = 'selected-column-table'
        #    elif col['name'] == switch.text and switch.value:
        #        col['classes'] = 'None'
        #        col['headerClasses'] = 'None'
        #    else:
        #        pass
        

    def flip_column_selection(col_name: str, switch_list: List[ui.switch], target_classes: str = 'None'):
        if KEY_CACHE.ctrl:
            switch_all(switch_list, False)

        col_obj: Dict = new_table.indexed_column_name_dict[col_name]
        classes_value = col_obj.get('classes', None)
        header_classes = col_obj.get('headerClasses', None)
        
        if (classes_value is None or classes_value == target_classes) and (header_classes is None or header_classes == target_classes):
            target_classes = 'selected-column-table'
        select_column(col_obj, target_classes)

    def index_columns():
        for col in new_table.columns:
            col['notselected'] = True
        new_table.indexed_column_name_dict = {col['name'] : col for col in new_table.columns}


    def select_column(col_obj: Dict, set_to_classes:str):

        col_obj['classes'] = set_to_classes
        col_obj['headerClasses'] = set_to_classes
        if set_to_classes != 'None':
            col_obj['notselected'] = False
        else:
            col_obj['notselected'] = True
        #new_table.update()
        column_selection_menu.update()

    def copy_selected(switch_list: List[ui.switch], table: CustomTable):
        
        selected = tuple(s.text for s in switch_list if s.value)
        sortedselected = sorted(table.selected, key=lambda x: x['#'])
        selected_row = tuple(tuple(item[key] if item[key] is not None else '' for key in selected) for item in sortedselected)
        finished = M.singlecopy('\n'.join(tuple('\t'.join(line) for line in selected_row)))
        if finished:
            notification('positive', '選択したセルをコピーしました。')
        else:
            notification('negative', '選択中の行はありません。')

    def switch_all(switch_list: List[ui.switch], switch_value: bool):
        for sw in switch_list:
            sw.set_value(switch_value)

    selected_col_dict = data.columns.to_list()
    switch_list: List[ui.switch] = []
    with ui.column().classes('w-full flex justify-end gap-2 mb-2') as container:
        with ui.row().classes('w-full'):
            with ui.button(icon='menu'):
                with ui.menu(), ui.column().classes('gap-0 p-2') as column_selection_menu:
                    for column in selected_col_dict:
                        if column != '#':
                            new_switch = ui.switch(column, value=True)#.bind_value(selected_col_dict, column)
                            new_switch.on_value_change(lambda n = new_switch: update_table(n))
                            switch_list.append(new_switch)
                    ui.button('全選択', on_click=lambda swl= switch_list, v=True: switch_all(swl, v)).classes('pl-2').props('dense flat')
                    ui.button('全取消', on_click=lambda swl= switch_list, v=False: switch_all(swl, v)).classes('pl-2').props('dense flat')
            ui.space()
            copy_chip = ui.chip(icon='content_copy', color='green', text='コピー')
            
        if isinstance(data, pd.DataFrame):
            data['#'] = range(1, len(data) + 1)
            new_table = CustomTable.from_pandas(data, row_key='#', selection='multiple').classes('w-full')

            for col in new_table.columns:
                if col['name'] == '#':
                    col['headerClasses'] = 'hidden'
                    col['classes'] = 'hidden'

        elif isinstance(data, list):
            new_table = CustomTable(data).classes('w-full')
        else:
            raise ValueError("Unsupported data type for table creation.")
    index_columns()
    for sw in switch_list:
        sw.bind_value_from(new_table.indexed_column_name_dict[sw.text], 'notselected')
    copy_chip.on_click(lambda: copy_selected(switch_list, new_table))
    
    new_table.add_slot('body-cell', r'''
                        <q-td :props="props" @dblclick="$parent.$emit('cell-click', props)">
                            {{ props.value }}
                        </q-td>
                    ''')
    new_table.on('cell-click', lambda e: flip_column_selection(col_name= e.args['col']['name'], switch_list=switch_list))
    with container:
        with ui.context_menu() as context_menu:
            context_copy_menu = ui.menu_item('選択行をコピー', on_click=lambda: copy_selected(switch_list, new_table))
            ui.separator()
            async def _table_excel_write():
                selected = tuple(s.text for s in switch_list if s.value)
                sortedselected = sorted(new_table.selected, key=lambda x: x['#'])
                select_row = [{f'Col_{index}' : item[key] if item[key] is not None else '' for index, key in enumerate(selected)} for item in sortedselected]
                #selected_row = tuple(tuple(item[key] if item[key] is not None else '' for key in selected) for item in sortedselected)
                await excel_write_dialog(select_row)
            ui.menu_item('Excelへの書き出し', on_click=_table_excel_write)

    new_table.context_menu['copy'] = context_copy_menu
    new_table.style('user-select: none; -webkit-user-select: none; -moz-user-select: none')
    
    #new_table.context_menu['copy'].on_click(lambda: copy_selected(switch_list, new_table))

    return container, new_table

def create_schedule_table(
    data: 'pd.DataFrame',
    event_cache: dict,
):
    """
    Schedule-specific table for 作業計画 screen.

    リスケ依頼生成 button is always visible. Clicking it opens a dropdown
    listing every process row with a checkbox and datepicker. Setting any
    datepicker cascades workday-offset dates to ALL subsequent rows (checked
    or not) and updates the request textarea. Only checked rows appear in the
    request string. Closing the dropdown resets the button text.
    """
    # ── dropdown state ────────────────────────────────────────────────────────
    dropdown_rows: list[dict] = []
    # Each entry: {'name': str, 'value': str}
    # value = YYYY-MM-DD string, '' means not set

    def _init_dropdown_rows() -> None:
        """Populate dropdown_rows from current table rows, all values blank.

        NOTE: sched_table is assigned later in the construction block; only
        call this after the UI has been built (e.g. from _on_menu_show).
        """
        nonlocal dropdown_rows
        dropdown_rows = [
            {'name': r['現在工程'], 'value': '', 'checked': True}
            for r in sorted(sched_table.rows, key=lambda x: x['#'])
        ]

    def _get_request_text() -> str:
        built_str = SU._build_request_text_from_list(dropdown_rows)
        if not built_str:
            raise ValueError('リスケ依頼の内容が空です。日付を設定してください。')
        return built_str        

    async def _write_to_excel() -> None:
        book_name = RESCHEDULE_EXCEL_CONFIG['path'].split('/')[-1].removesuffix('?web=1')

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

        active_task: asyncio.Task | None = None
        exited = False

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
            active_task = asyncio.create_task(append_reschedule_row(
                request_str=_get_request_text(),
                username=USER_INS.user_name or USER_INS.uid,
                lotid=event_cache.get('lotid', ''),
                db_figure=event_cache.get('productcode', ''),
                file_path=RESCHEDULE_EXCEL_CONFIG['path'],
                sheet_name=RESCHEDULE_EXCEL_CONFIG['sheet'],
            ))
            await active_task

            if not exited:
                status_icon.props('name=check_circle color=positive size=xl')
                status_label.set_text('書き込み完了 — Excelを閉じると続行できます')
                active_task = asyncio.create_task(wait_for_book_close(book_name))
                await active_task
                notification('positive', '書き出し完了')

        except asyncio.CancelledError:
            pass
        except Exception as exc:
            notification('negative', f'書き出し失敗: {exc}')
        finally:
            reveal_task.cancel()
            overlay.close()
            overlay.delete()

    def _rebuild_dropdown_ui() -> None:
        """Clear and redraw the dropdown column from current dropdown_rows state.

        Each row has a checkbox to include/exclude it from the request string.
        Dates are always calculated for all rows regardless of checked state;
        only checked rows appear in the output text.
        Holds live inp references so date cascade updates widgets directly
        (no recursive rebuild on each date change).
        """
        reschedule_menu_col.clear()
        with reschedule_menu_col:
            if not dropdown_rows:
                ui.label('行がありません。').classes('text-sm p-2')
                return

            cb_list: list = []
            inp_list: list = []

            for i, row_data in enumerate(dropdown_rows):
                with ui.row(align_items='center').classes('w-full gap-2 px-2 py-1'):
                    cb = ui.checkbox(value=row_data['checked'])
                    cb_list.append(cb)
                    ui.label(row_data['name']).classes('text-sm font-bold flex-1 min-w-0 truncate')
                    inp = ui.input(value=row_data['value']).props('type=date dense outlined')
                    inp_list.append(inp)

            def _get_text() -> str:
                return SU._build_request_text_from_list(dropdown_rows)

            preview = ui.textarea(value=_get_text()).props(
                'outlined dense rows=4'
            ).classes('w-full px-2 mt-1')

            def _make_checkbox_handler(i: int):
                def _on_check(e) -> None:
                    dropdown_rows[i]['checked'] = e.value
                    preview.value = _get_text()
                return _on_check

            def _make_inp_date_handler(i: int):
                def _on_date_change(e) -> None:
                    new_val = e.value if e.value is not None else ''
                    dropdown_rows[i]['value'] = new_val
                    try:
                        base = date.fromisoformat(new_val)
                        subsequent_names = [
                            dropdown_rows[k]['name'] for k in range(i + 1, len(dropdown_rows))
                        ]
                        cascaded_dates = SU.cascade_dates_with_capacity(
                            trigger_name=dropdown_rows[i]['name'],
                            trigger_date=base,
                            subsequent_names=subsequent_names,
                        )
                        for j, (subseq_inp, cascaded) in enumerate(zip(inp_list[i + 1:], cascaded_dates)):
                            iso = cascaded.isoformat()
                            dropdown_rows[i + 1 + j]['value'] = iso
                            subseq_inp.value = iso
                    except (ValueError, TypeError):
                        pass
                    preview.value = _get_text()
                return _on_date_change

            for j, cb in enumerate(cb_list):
                cb.on_value_change(_make_checkbox_handler(j))

            for j, inp in enumerate(inp_list):
                inp.on_value_change(_make_inp_date_handler(j))

            # Action chips
            with ui.row().classes('gap-2 px-2 py-2 w-full'):
                ui.chip(
                    'コピー', icon='content_copy', color='green',
                    on_click=lambda: M.singlecopy(_get_text()),
                )
                ui.chip('ファイルに書き出す', icon='upload_file',
                        on_click=_write_to_excel)

    # ── row setup ─────────────────────────────────────────────────────────────
    data = data.copy()
    data['#'] = range(1, len(data) + 1)

    switch_list: list = []
    selected_col_dict = data.columns.to_list()

    with ui.column().classes('w-full flex justify-end gap-2 mb-2') as container:
        with ui.row().classes('w-full'):
            with ui.button(icon='menu'):
                with ui.menu(), ui.column().classes('gap-0 p-2') as column_selection_menu:
                    for column in selected_col_dict:
                        if column != '#':
                            sw = ui.switch(column, value=True)
                            switch_list.append(sw)
                    ui.button(
                        '全選択',
                        on_click=lambda swl=switch_list: [s.set_value(True) for s in swl],
                    ).classes('pl-2').props('dense flat')
                    ui.button(
                        '全取消',
                        on_click=lambda swl=switch_list: [s.set_value(False) for s in swl],
                    ).classes('pl-2').props('dense flat')

            ui.space()
            ui.chip(icon='content_copy', color='green', text='コピー',
                    on_click=lambda: _copy_selected())

            # Always-visible reschedule button — no set_visibility(False)
            reschedule_btn = ui.chip('リスケ依頼生成', icon='edit_note', on_click= lambda: None)
            with reschedule_btn:
                with ui.menu() as reschedule_menu:
                    with ui.column().classes('p-2 min-w-96') as reschedule_menu_col:
                        pass  # filled by _rebuild_dropdown_ui on @show

        sched_table = CustomTable.from_pandas(data, row_key='#',
                                              selection='multiple').classes('w-full')

        for col in sched_table.columns:
            if col['name'] == '#':
                col['headerClasses'] = 'hidden'
                col['classes'] = 'hidden'

    # ── button toggle via menu show/hide events ───────────────────────────────
    def _on_menu_show(_=None) -> None:
        reschedule_btn.set_text('取消')
        _init_dropdown_rows()
        _rebuild_dropdown_ui()

    def _on_menu_hide(_=None) -> None:
        reschedule_btn.set_text('リスケ依頼生成')

    reschedule_menu.on('before-show', _on_menu_show)
    reschedule_menu.on('before-hide', _on_menu_hide)

    # ── column visibility bindings ────────────────────────────────────────────
    sched_table.indexed_column_name_dict = {col['name']: col for col in sched_table.columns}
    for sw in switch_list:
        if sw.text in sched_table.indexed_column_name_dict:
            col_obj = sched_table.indexed_column_name_dict[sw.text]
            col_obj['notselected'] = True
            sw.bind_value_from(col_obj, 'notselected')
            sw.on_value_change(lambda n=sw: _toggle_col(n))

    # ── copy helper ───────────────────────────────────────────────────────────
    def _copy_selected():
        selected = tuple(s.text for s in switch_list if s.value)
        sorted_sel = sorted(sched_table.selected, key=lambda x: x['#'])
        rows_data = tuple(
            tuple(item[k] if item[k] is not None else '' for k in selected)
            for item in sorted_sel
        )
        finished = M.singlecopy('\n'.join('\t'.join(str(c) for c in line) for line in rows_data))
        if finished:
            notification('positive', '選択したセルをコピーしました。')
        else:
            notification('negative', '選択中の行はありません。')

    def _toggle_col(switch: ui.switch):
        col_obj = sched_table.indexed_column_name_dict[switch.text]
        col_obj['classes'] = 'None' if switch.value else 'selected-column-table'
        col_obj['headerClasses'] = 'None' if switch.value else 'selected-column-table'
        col_obj['notselected'] = switch.value
        column_selection_menu.update()

    # ── double-click column highlight (same pattern as create_custom_table) ──
    def _switch_all(value: bool) -> None:
        for sw in switch_list:
            sw.set_value(value)

    def _select_col(col_obj: dict, set_to_classes: str) -> None:
        col_obj['classes'] = set_to_classes
        col_obj['headerClasses'] = set_to_classes
        col_obj['notselected'] = (set_to_classes == 'None')
        column_selection_menu.update()

    def _flip_col(col_name: str) -> None:
        if KEY_CACHE.ctrl:
            _switch_all(False)
        col_obj = sched_table.indexed_column_name_dict.get(col_name)
        if col_obj is None:
            return
        classes_value = col_obj.get('classes', None)
        header_classes = col_obj.get('headerClasses', None)
        target_classes = 'None'
        if (classes_value is None or classes_value == target_classes) and \
                (header_classes is None or header_classes == target_classes):
            target_classes = 'selected-column-table'
        _select_col(col_obj, target_classes)

    sched_table.add_slot('body-cell', r'''
        <q-td :props="props" @dblclick="$parent.$emit('cell-click', props)">
            {{ props.value }}
        </q-td>
    ''')
    sched_table.on('cell-click', lambda e: _flip_col(e.args['col']['name']))

    sched_table.style(
        'user-select: none; -webkit-user-select: none; -moz-user-select: none'
    )

    return container, sched_table

async def schedule_quick_peek(tabs_obj: ui.tabs, tab_obj: ui.tab, schedule_: Dict[str, pd.DataFrame], event_cache) -> Tuple[pd.DataFrame, ui.table]:
    with ui.tab_panel(tab_obj).classes('w-full'):
        # Process History content (initially hidden)
        with ui.column().classes('w-full') as history_content:
            if schedule_ is None:
                schedule_, _ = await query_data_stdin('v', event_cache['lotid'])
            else:
                tabs_obj.set_value(tab_obj)
            ui.label('作業計画').classes('text-lg mb-2')
            df_schedule = None
            if not schedule_:
                ui.label('作業計画はありません。').classes('text-lg mb-2')
                scheduletable = None
                df_schedule = None
            else:
                #for k in schedule_.keys():
                #    df_schedule = schedule_[k]
                    #scheduletable = ui.table.from_pandas(df_schedule).classes('w-full')
                #    _, scheduletable = create_custom_table(df_schedule)

                for k in schedule_.values():
                    df_schedule = k
                    #scheduletable = ui.table.from_pandas(df_schedule).classes('w-full')
                    _, scheduletable = create_schedule_table(df_schedule, event_cache)
                    
    return df_schedule, scheduletable

async def history_quick_peek(tabs_obj: ui.tabs, tab_obj: ui.tab, history_: Dict[str, pd.DataFrame], event_cache) -> Tuple[pd.DataFrame, ui.table]:
    with ui.tab_panel(tab_obj).classes('w-full overscroll-auto'):    
        if history_ is None:
            history_, _ = await query_data_stdin('h', event_cache['lotid'])
        else:
            tabs_obj.set_value(tab_obj)
        # Lot Information content (initially hidden)
        with ui.column().classes('w-full') as lot_content:
            ui.label('処理履歴').classes('text-lg mb-2')
            if not history_:
                ui.label('作業履歴はありません。').classes('text-lg mb-2')
                historytable = None
                df_lot = None
            else:
                #for k in history_.keys():
                #    df_lot = history_[k]
                #    _, historytable = create_custom_table(df_lot)
                
                for k in history_.values():
                    df_lot = k
                    _, historytable = create_custom_table(df_lot)
    return df_lot, historytable

async def restriction_qucik_peek(tabs_obj: ui.tabs, tab_obj: ui.tab, restriction_: Dict[str, pd.DataFrame], event_cache) -> Tuple[pd.DataFrame, ui.table]:
    with ui.tab_panel(tab_obj).classes('w-full'):
        # Additional Data content (initially hidden)
        with ui.column().classes('w-full') as additional_content:
            if restriction_ is None:
                restriction_, _ = await query_data_stdin('y', event_cache['lotid'])
            else:
                tabs_obj.set_value(tab_obj)
            ui.label('出荷制限情報').classes('text-lg mb-2')
            if not restriction_:
                ui.label('出荷制限情報はありません。').classes('text-lg mb-2')
                restriction_table = None
                df_restriction = None
            else:
                #for k in restriction_.keys():
                    #df_restriction = pd.DataFrame(restriction_[k])
                #    df_restriction = restriction_[k]
                #    _, restriction_table = create_custom_table(df_restriction)
                
                for k in restriction_.values():
                    #df_restriction = pd.DataFrame(restriction_[k])
                    df_restriction = k
                    _, restriction_table = create_custom_table(df_restriction)
    return df_restriction, restriction_table

async def serial_quick_peek(tabs_obj: ui.tabs, tab_obj: ui.tab, serial_: Dict[str, pd.DataFrame], event_cache) -> Tuple[pd.DataFrame, ui.table]:
    with ui.tab_panel(tab_obj).classes('w-full'):
        # Process History content (initially hidden)
        with ui.column().classes('w-full') as history_content:
            if serial_ is None:
                serial_, _ = await query_data_stdin('s', event_cache['lotid'])
            else:
                tabs_obj.set_value(tab_obj)
            ui.label('シリアル').classes('text-lg mb-2')
            df_serial = None
            if not serial_:
                ui.label('シリアルデータはありません。').classes('text-lg mb-2')
                serialtable = None
                df_serial = None
            else:
                #for k in serial_.keys():
                #    df_serial = serial_[k]
                #    _, serialtable = create_custom_table(df_serial)
                
                for k in serial_.values():
                    df_serial = k
                    _, serialtable = create_custom_table(df_serial)
    return df_serial, serialtable


async def create_process_dialog(history_: Dict[str, pd.DataFrame] = None, schedule_: Dict[str, pd.DataFrame] = None, restriction_: Dict[str, pd.DataFrame] = None, serial_: Dict[str, pd.DataFrame] = None):
    #print('create_process_dialog')
    # Create dialog with proper sizing
    
    async def lazyload(event: events.ValueChangeEventArguments):
        
        spinner_open()
        current = event.value
        if dialog.inited_tab.get(current) is not None or current == '指図詳細':
            spinner_close()
            return
        func = None
        for f in process_tab_dict.values():
            if f[0] == current:
                func = f[1]
                break
        
        tab_dict = {
            '作業実績' : history,
            '作業予定' : schedule,
            '出荷制限' : restriction,
            'シリアル' : serial
        }
        
        tab = tab_dict.get(current)
        

        with t_panels:
            _, _ = await func(process_tabs, tab, None, event_cache)
        
        dialog.inited_tab[current] = 1
        spinner_close()
    
    try:
        event_cache:Dict[str, str] = CLICK_EVENT_CACHE.args['data']
    except Exception as e:
        notification('negative', f'予想外のエラーが発生しました,リトライしてください。：{e}')
        spinner_close()
        return

    process_tab_dict = {
        'h' : ('作業実績', partial(history_quick_peek)),
        'p' : ('作業予定', partial(schedule_quick_peek)),
        'r' : ('出荷制限', partial(restriction_qucik_peek)),
        's' : ('シリアル', partial(serial_quick_peek))
    }

    if any(parm is not None for parm in (history_, schedule_, restriction_, serial_)):
        if history_ is not None:
            flag = 'h'
            data = history_
        if schedule_ is not None:
            flag = 'p'
            data = schedule_
        if restriction_ is not None:
            flag = 'r'
            data = restriction_
        if serial_ is not None:
            flag = 's'
            data = serial_

        name_func = process_tab_dict[flag]
        name = name_func[0]
        func = name_func[1]
        
        with ui.context.client.content:
            with ui.dialog().classes('w-full h-5/6').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:
                close_btn = ui.label('閉じる').classes('text-xl text-white/75 hover:text-white cursor-pointer')
                close_btn.on('click', dialog.close)
                with ui.card().classes('w-full h-5/6').style('max-width: none'):
                    with ui.splitter(value=10, limits=[5, 10]).classes('w-full h-full') as splitter:
                        with splitter.before:
                            with ui.tabs().props('vertical').classes('w-full') as process_tabs:
                                quick_peek_tab = ui.tab(name)
                        with splitter.after:
                            with ui.tab_panels(process_tabs, value=quick_peek_tab).props('vertical').classes('w-full h-full'):
                                await func(process_tabs, quick_peek_tab, data, event_cache)
        spinner_close()
    else:
        location = await M.OC.asy_q_execute(
            '''
            SELECT 
                ZSLH.LOT_ID, ZSLH.SHELF_NO, ZSLH.IN_OUT_FLG, ZMMS.SHELF_NAME 
                FROM Z_SHELF_LOT_HISTORY ZSLH LEFT JOIN Z_MES_MASTER_SHELF_LC ZMMS ON ZSLH.SHELF_NO = ZMMS.SHELF_NO 
                WHERE LOT_ID = :1    
        ''', (event_cache['lotid'],))

        if location:
            last_seen = location[-1]
        else:
            last_seen = [None, None, None, None]
        schedule_info: List[Tuple[int, datetime]] = await M.OC.asy_q_execute(
            '''
            SELECT 
                (SELECT COUNT(*) FROM LV23_NHK.V_ASP_生産計画一覧 WHERE 指図NO = :1) as count,
                V_ASP_生産計画一覧.完了予定日 
            FROM LV23_NHK.V_ASP_生産計画一覧 V_ASP_生産計画一覧 
            WHERE V_ASP_生産計画一覧.指図NO = :1 
            ORDER BY V_ASP_生産計画一覧.着手予定日 DESC 
            FETCH FIRST 1 ROWS ONLY    
        ''', (event_cache['lotid'], event_cache['lotid']))

        with ui.context.client.content:
            with CustomDialog().classes('w-full h-5/6').props('backdrop-filter="blur(8px) brightness(40%)"') as dialog:
                close_btn = ui.label('閉じる').classes('text-xl text-white/75 hover:text-white cursor-pointer')
                close_btn.on('click', dialog.close)
                with ui.card().classes('w-full h-5/6').style('max-width: none'):
                    with ui.splitter(value=10).classes('w-full h-full') as splitter:
                        with splitter.before:
                            with ui.tabs().props('vertical').classes('w-full') as process_tabs:
                                details = ui.tab('指図詳細')#, icon='mail')
                                schedule = ui.tab('作業予定')
                                history = ui.tab('作業実績')
                                restriction = ui.tab('出荷制限')
                                serial = ui.tab('シリアル')
                            with splitter.after:
                                with ui.tab_panels(process_tabs, value=details).props('vertical').classes('w-full h-full') as t_panels:
                                    #df_schedule, scheduletable = await schedule_quick_peek(process_tabs, schedule, schedule_, event_cache)
                                    #df_lot, historytable = await history_quick_peek(process_tabs, history, history_, event_cache)
                                    #df_restriction, restriction_table = await restriction_qucik_peek(process_tabs, restriction, restriction_, event_cache)
                                    #df_serial, serialtable = await serial_quick_peek(process_tabs, serial, serial_, event_cache
                                    with ui.tab_panel(details).classes('bg-slate-50'):
                                        with ui.column().classes('bg-gray-50 rounded-lg shadow-md p-6 max-w-full text-gray-800'):
                                            # Dashboard title
                                            ui.label(f'指図詳細  {event_cache["lotid"]}').classes('text-2xl font-semibold mb-6 pb-3 border-b-2 border-gray-200 text-gray-800')
                                            
                                            # Dashboard grid
                                            with ui.row().classes('grid grid-cols-2 gap-4 w-full'):
                                                # Column 1 - 3 items
                                                with ui.column().classes('grid auto-rows-auto gap-2 content-start'):
                                                    # Item 1 - 出荷制限
                                                    with ui.card().classes('bg-white rounded-lg p-4 shadow-lg transition-all duration-200 ease-in-out hover:-translate-y-0.5 hover:shadow-lg') as restriction_tile:
                                                        ui.label('出荷制限').classes('text-sm font-medium text-gray-500 mb-2')
                                                        restriction_text = '出荷制限中' if event_cache['restriction'] != '-' else '出荷可能'
                                                        restriction_color = 'text-red-600' if event_cache['restriction'] != '-' else 'text-blue-600'
                                                        ui.label(restriction_text).classes(f'text-2xl font-bold {restriction_color}')
                                                    
                                                    # Item 2 - 残工程数
                                                    with ui.card().classes('bg-white rounded-lg p-4 shadow-lg transition-all duration-200 ease-in-out hover:-translate-y-0.5 hover:shadow-lg') as schedule_tile:
                                                        ui.label('残工程数').classes('text-sm font-medium text-gray-500 mb-2')
                                                        on_hold = event_cache.get('onholdreason', None)
                                                        process_count = str(schedule_info[0][0]) if schedule_info and on_hold == '-' else '保留中' if on_hold != '-' else '0'
                                                        process_color = 'red' if process_count == '保留中' else 'blue'
                                                        ui.label(process_count).classes(f'text-2xl font-bold text-{process_color}-600')
                                                        ui.separator().classes('w-full')
                                                        ui.label('完成予定日').classes('text-sm font-medium text-gray-500 mb-2')
                                                        completion_date = schedule_info[0][1].strftime("%Y-%m-%d %H:%M:%S") if schedule_info else '完成予定はありません'
                                                        ui.label(completion_date).classes('text-2xl font-bold text-blue-600')
                                                    
                                                    # Item 3 - 完成予定日
                                                    #with ui.card().classes('bg-white rounded-lg p-4 shadow-sm transition-all duration-200 ease-in-out hover:-translate-y-0.5 hover:shadow-lg'):
                                                        
                                                
                                                # Column 2 - 1 item with multiple values
                                                with ui.column().classes('grid grid-rows-2 gap-4'):
                                                    with ui.card().classes('bg-white rounded-lg p-4 shadow-lg transition-all duration-200 ease-in-out hover:-translate-y-0.5 hover:shadow-lg'):
                                                        ui.label('倉庫棚番').classes('text-sm font-medium text-gray-500 mb-2')
                                                        shelf_location = "倉庫預かり中" if last_seen[1] is None else f'{last_seen[1]}-{last_seen[3]}'
                                                        ui.label(shelf_location).classes('text-2xl font-bold text-blue-600')
                                                        
                                                        status_map = {1: "エリア内", 2: "搬出済", 0: "エリア外"}
                                                        status_text = status_map.get(last_seen[2], "状態不明")
                                                        ui.label(status_text).classes('text-2xl font-bold text-blue-600')

                                            

                                    t_panels.on_value_change(lazyload)           

    #spinner_close()

    dialog.open()
    
    return dialog




async def process_lot(input_obj: ui.input | ui.textarea, shelf: ui.input | ui.textarea, table_obj: ui.table, mode: str):
    raw_input = input_obj.value
    raw_shelf = shelf.value
    if not raw_input:
        notification('warning', '指図の入力がありません。')
        return
    if not raw_shelf and mode == 'IN':
        #notification('warning', '棚番の入力がありません。')
        return
    existing_lot = [obj['lotid'] for obj in table_obj.rows]
    process_data = await uni_query_tp(raw_input, existing_lot)
    #print(f'process data = {process_data}')
    if not process_data:
        return
    #placeholder ==================
    procode = process_data.pop('procode', None)
    if procode is None:
        notification('info', f'{process_data['lotid']}の工程コード情報に異常があります、そのまま投入することはできますが、履歴ではエラーと表示します。')

    if mode == 'IN':
        shelf_validator = {'process': 'AA'} #need a generator for validator or maybe store it somewhere
        correct_shelf = shelf_validator['process'] #need a disposable data (maybe a resouce data column) for shelf validation
        if correct_shelf != shelf:
            result = await quick_dialog_yes_no('投入しようとしている棚番は予定の棚番ではありません、間違いではありませんか？\nそのまま投入する場合は「はい」を押してください。\nキャンセルする場合は「いいえ」を押してください。')
            if result == 1:
                process_data['shelf'] = raw_shelf
                table_obj.add_row(process_data)
                input_obj.set_value(None)
                notification('positive', f'棚番 {raw_shelf} に指図 {raw_input} を投入しました。')
                return process_data
            else:
                notification('negative', f'棚番 {raw_shelf} への投入をキャンセルしました。')
                return None
    else:
        if not process_data['shelf']:
            notification('warning', '搬出しようとしている指図は棚番が登録されていません、棚番転記の操作によって状態の変化はありません。', postion='center')
        table_obj.add_row(process_data)
        input_obj.set_value(None)
        notification('positive', f'棚番 {'外' if process_data['shelf'] is None else process_data['shelf']} から指図 {raw_input} を搬出しました。')
        return process_data



async def switch_mode(button_obj: ui.button, container:ui.card, shelf_input:ui.input, lotinput:ui.input, table: ui.table, lock_bt: ui.button):
    #print(not shelf_input.value, shelf_input.value)
    
    if not table.rows and not lotinput.value:
        if 'IN' in button_obj.text:
            text = 'OUT'
            color = 'red'
            _icon = 'north'
        else:
            text = 'IN'
            color = 'green'
            _icon = 'south'        
        button_obj.text = text
        #print(button_obj._props, button_obj._style)
        #button_obj.props(f'color={{"{color}"}}')
        #button_obj.props(remove='color icon')
        button_obj.props(remove='color icon', add=f'color={color} icon={_icon}')
        container.classes(replace=f'h-full w-full bg-{color}-50 p-4 space-y-4')

        shelf_input.enable()


        lotinput.set_value(None)
        toggle_input(shelf_input, lock_bt, mode=text)

        button_obj.update()
    elif lotinput.value:
        text = button_obj.text
        if text == 'IN' and not shelf_input.value:
            notification('info', message='棚番が入力されていません。', postion='center')
            return
        if shelf_input.enabled:
            toggle_input(shelf_input, lock_bt, mode=text)
        await process_lot(lotinput, shelf_input, table, mode=text)
    elif not lotinput.value and table.rows:
        notification('info', multiline=True, message='すでに入力されたデータがあります、\nモードを変更するには先に入力済みのデータを登録してください。', postion='center')


async def register(modeswitch: ui.button, table: ui.table):
    pooled = table.rows
    row_count = len(pooled)
    current_mode = modeswitch.text
    parameters = [(row['shelf'] if current_mode == 'IN' else '', '1' if current_mode == 'IN' else '0', USER_INS.uid, row['lotid']) for row in pooled]
    await M.modify(parameters, 'U')
    parameters_insert = [(row['lotid'], row['shelf'], '1' if current_mode == 'IN' else '0', USER_INS.uid, 0) for row in pooled]
    await M.modify(parameters_insert, 'I')
    #check_result = await M.OC.sql_plain(
    #    '''
    #    SELECT LOT_ID, SHELF_NO, USER_CODE, IN_OUT_FLG FROM Z_SHELF_LOT_HISTORY WHERE LOT_ID IN ('4232045315', '4232045317')    
    #'''
    #)
    clear_table(table)
    #table.update()
    notification('positive', f'{row_count}件を転記しました。')
    #print(check_result)
    

def toggle_input(_input: ui.input, button: ui.button, mode: str, direct_unlock: bool = False):
    def enable():
        _input.enable()
        button.props(remove='label', add='label="確定"')
        button.props(remove='color', add='color="primary"')

    def disable(label: str = "解除"):
        _input.disable()
        button.props(remove='label', add=f'label="{label}"')
        button.props(remove='color', add='color="red"')

    def refresh():
        _input.set_value(None)

    if direct_unlock:
        if _input.enabled:
            if not _input.value:
                notification('info', message='棚番が入力されていません。', postion='center')
                return
            disable()
            return
        else:

            refresh()
            enable()
            return

    if mode == 'OUT':

        refresh()
        label = 'OUT'
        button.disable()
        disable(label)
        return

    if mode == 'IN' and button.enabled:
        
        if not _input.value:
            notification('info', message='棚番が入力されていません。', postion='center')
            return
        else:
            disable()
            #refresh()
            #enable()
            return
            

    if mode == 'IN' and not button.enabled:

        refresh()
        button.enable()
        enable()
        return

    disable()




def clear_table(table: ui.table):
    table.rows.clear()
    table.update()

class TPSystem:
    def __init__(self):
        self.tpmanger : bool = False
        self.transport : bool = False
        self.tpmanger_bt : ui.item = None
        self.transport_bt: ui.item = None
        self.tpmanger_instance: ui.card = None
        self.transport_instance: ui.card = None
        self.manager_ag: ui.aggrid = None

    def init_instance(self):
        self.transport_instance = transport_system()
        container, manager_ag = shelf_viewer()
        self.tpmanger_instance = container
        self.manager_ag = manager_ag
        #self.transport_instance.set_visibility(True)
        #self.tpmanger_instance.set_visibility(False)
        self.tpmanger = True
        self.transport = False
        self.transport_instance.bind_visibility(self, 'tpmanger')
        self.tpmanger_instance.bind_visibility(self, 'transport')

    def set_caller(self, manager_bt: ui.item, transport_bt: ui.item):
        self.tpmanger_bt = manager_bt
        self.transport_bt = transport_bt
        self.tpmanger_bt.bind_enabled(self, 'tpmanger')
        self.transport_bt.bind_enabled(self, 'transport')

    def call_transport(self) -> ui.card:
        self.tpmanger = True
        self.transport = False
        self.transport_instance.set_visibility(True)
        return self.transport_instance
    
    async def call_tpmanger(self) -> ui.card:
        self.transport = True
        self.tpmanger = False
        self.tpmanger_instance.set_visibility(True)
        try:
            df = await tp_query()
            self.manager_ag.from_pandas(df)
        except E.NoDataExc:
            notification('negative', '取得できるデータはありませんでした。')
            notification('ongoing','取得できるデータがない不具合：１．インターネットの接続を確認 ２.社内サーバーの状態を確認', postion='center', close_button=True, color='black')
        except Exception as e:
            notification('negative', f'エラーが発生しました：　　{e}')
        finally:
            spinner_close()
        return self.tpmanger_instance


def transport_system() -> ui.card:
    def clear_input():
        shelf.set_value(None)
        lot.set_value(None)

    with ui.card().classes('h-full w-full').classes('bg-sky-200') as main_container:
        with ui.row().classes('w-full text-left') as header_container:
            ui.label('棚番管理').classes('text-lg')
        with ui.card().classes('h-full w-full bg-green-50') as result_container:
            with ui.column(align_items='start').classes('w-full') as input_row:
                #in_bt = ui.button('IN')
                #ui.space()
                with ui.row(align_items='start').classes('w-full'):
                    shelf = ui.input(placeholder='棚番').props('square standout').classes('text-lg w-1/4')
                    lock = ui.button('確定').props('size=lg')

                with ui.row(align_items='start').classes('w-full'):
                    lot = ui.input(placeholder='指図番号').props('square standout').classes('text-lg w-1/2')
                    ui.space()
                    lotclear = ui.button('入力データクリア', on_click=clear_input, color='red').props('size=lg')
                    lock.on_click(lambda ip=shelf, bt=lock: toggle_input(ip, bt, mode_switch.text, True))
                

                #ui.space()
                #out_bt = ui.button('OUT')
            with ui.row().classes('w-full item-left') as status_row:
                mode_switch = ui.button('IN', icon='south').classes('w-full').props('color=green size=xl')
                
                #status_label = ui.label('状態:').classes('text-lg')
                #status = ui.label('-').classes('text-lg')
                ui.separator()
        pool_column = [{'name': 'shelf', 'label': '棚番', 'field': 'shelf'},
                        {'name': 'lotid', 'label': '指図番号', 'field': 'lotid'},
                        {'name': 'process', 'label': '工程', 'field': 'process'},
                        {'name': 'to_area', 'label': '投入エリアヒント', 'field': 'to_area'},
                        {'name': 'status', 'label': '状態', 'field': 'status'},
                        ]
        #placeholder = [{'shelf': 'データがありません', 'lotid': 'INまたはOUTを押してデータを追加してください。', 'process': '-', 'status': '-'}]    
        with ui.row().classes('w-full'):
            commit = ui.button('データ登録', color='green').props('size=lg')
            ui.space() 
            clear = ui.button('全データクリア', color='red').props('size=lg')  
        pool_table = ui.table(
            columns=pool_column,
            row_key='lotid',
            rows=[]
        ).classes('w-full h-full text-lg')
            

        mode_switch.on_click(lambda ms = mode_switch, con = result_container, she=shelf, lo=lot, tb=pool_table, lb=lock: switch_mode(button_obj=ms, container=con, shelf_input=she, lotinput=lo, table=tb, lock_bt=lb))
        #in_bt.on_click(lambda ip=lot, sh=shelf, tb=pool_table, st=mode_switch.text, ipt='IN' : process_lot(ip, sh, tb, st, ipt))
        commit.on_click(lambda sw=mode_switch, tb= pool_table: register(sw, tb))    
        clear.on_click(lambda t = pool_table: clear_table(t))


    return main_container

def shelf_viewer() -> ui.card:
    #df = await tp_query()
    df = pd.DataFrame()
    ag = AgGrid()
    ag.init_column_settings('A',False)
    
    with ui.card().classes('h-full w-full') as container:
        ui.label('棚番一覧').classes('text-lg')
        with ui.row().classes('w-full bg-zinc-300 rounded-lg p-2 shadow-sm transition-all duration-200 ease-in-out hover:-translate-y-0.5 hover:shadow-lg') as top_container:
            wrong_shelf = ui.chip('エリア異常', icon='question_mark', color='red')
            urgent = ui.chip('急ぎ', icon='priority_high', color='orange')
            caption = ui.label('ステータスモニター').classes('text-2xl font-semibold text-gray-400')
            wrong_shelf.set_visibility(False)
            urgent.set_visibility(False)
        tp_all_grid = ag.aggrid_creator(df)
    ag.cellkey_func_bind()
    
    return container, tp_all_grid


@ui.page('/tp')
async def tp_app(client: Client):
    global USER_INS
    USER_INS = UserInstance()
    spinner_create()
    instance_manger = TPSystem()

    duplicated = await check_duplicated_page(client, TP_APP_NAME)
    
    if duplicated:
        return

    instance_manger.init_instance()
    with ui.left_drawer().classes('bg-blue-50').props('bordered') as tp_left_drawer:
        menu_title = ui.label('メニュー').classes('text-lg font-bold p-4')
        with ui.list().classes('w-full') as drawer_list:
            transport = ui.item('運搬システム', on_click=instance_manger.call_transport)
            tpmanagersys = ui.item('運搬状況一覧', on_click=instance_manger.call_tpmanger)
            ui.separator()
            mainsys = ui.item('プロセスマネージャー', on_click=lambda: ui.navigate.to('/'))
            ui.separator()
            about = ui.item('About', on_click=about_screen)
            ui.separator()
            ui.item('refresh', on_click=lambda: print('refresh'))
    with ui.header().classes('flex justify-between items-center bg-indigo-900 text-white p-4'):
        with ui.row().classes('w-full flex justify-between items-center'):
            ui.button('', icon='menu').props('flat').classes('text-white').on('click', lambda: tp_left_drawer.toggle())
            
            # Center - App name and version
            ui.label(f'{TP_APP_NAME} - v{TP_APP_VERSION}').classes('text-xl font-bold')
    
    instance_manger.set_caller(tpmanagersys, transport)
    
    initial_screen = instance_manger.call_transport()

    tp_left_drawer.hide()

    #tps_con = transport_system()
    #tpmanagersys.on_click(lambda: tps_con.set_visibility(False))


@ui.page('/transcation')
async def tp_app(client: Client):
    global USER_INS
    USER_INS = UserInstance()
    spinner_create()
    instance_manger = TPSystem()

    duplicated = await check_duplicated_page(client, 'PLACEHOLDER')
    
    if duplicated:
        return
    
    import transcation as trc

    trc.init_screen()


# Initialize the app


    
def attach_to_parent_console():
    if os.name != 'nt' or TEST:
        return False
    import ctypes
    # This is the correct way to access kernel32
    kernel32 = ctypes.windll.kernel32
    
    # AttachConsole(-1) attaches to parent process console
    # Returns 0 on failure, non-zero on success
    result = kernel32.AttachConsole(-1)
    
    if result:
        # Reopen stdout/stderr to console
        sys.stdout = open('CONOUT$', 'w')
        sys.stderr = open('CONOUT$', 'w')
        sys.stdin = open('CONIN$', 'r')
        return True
    
    # Get last error (optional debugging)
    error = kernel32.GetLastError()
    if error == 5:  # ERROR_ACCESS_DENIED
        # Already attached to a console
        return True
    
    return False

    

def main():
    #create_app()
    #app.on_startup(sigleton)
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    bin_path = os.path.join(desktop, 'processMgr')
    logging.basicConfig(
        level=logging.DEBUG,
        filename= os.path.join(bin_path, 'packgedversion_log.log'),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    logger = logging.getLogger(__name__)

    # Wire _app_logger to the same file handler so guitest errors are persisted
    if not _app_logger.handlers:
        _fh = logging.FileHandler(os.path.join(bin_path, 'packgedversion_log.log'), encoding='utf-8')
        _fh.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        _app_logger.addHandler(_fh)
        _app_logger.setLevel(logging.DEBUG)

    @app.exception_handler(Exception)
    async def handle_exception(request, exc):
        logger.exception(f"Unhandled exception: {exc}")
        ui.notify(f"Error occurred: {str(exc)}", type='negative', timeout=5000)
        return

    @app.on_startup
    async def close_splash():
        try:
            import pyi_splash
            pyi_splash.close()
        except ImportError:
            pass

    app.colors(primary=ACCENT_COLOR, secondary=SUB_ACCENT_COLOR)

    if frozen:
        newfeature = os.path.join(bin_path, 'new.txt')
        try:
            with open(newfeature, 'r') as newf:
                newf.seek(0)
                version = newf.read()
                if version == '':
                    version = '0.0.0'
        except FileNotFoundError:
            version = '0.0.0'

        if int(version.replace('.', '')) < int(VERSION_NUMBER_MAIN.replace('.', '')):
            try:
                import webbrowser
                webbrowser.open(r"\\sflise02\伊勢原２\接合・セラミック部\業務課\2.業務G\Apps\プロセスマネージャー\newsletter.html")
                with open(newfeature, 'w+') as f:
                    f.write(VERSION_NUMBER_MAIN)
            except Exception as e:
                notification('negative', f'アップデート詳細取得時にエラーが発生しました：{e}')

        try:
            ui.run(title=APP_NAME, favicon='🏭', language='ja', reload= False, native= TEST, storage_secret='rosesarerosie', reconnect_timeout=30, port=54322, window_size=(1024,768))
        except Exception as e:
            logger.exception(f"Main runner exception: {e}")
    else:
        try:
            ui.run(title=APP_NAME, favicon='🏭', language='ja', reload= False, storage_secret='rosesarerosie', port=54321, native=TEST, reconnect_timeout=30) #native= not TEST, port=54321, window_size=(1280, 800))
        except Exception as e:
            logger.exception(f"Main runner exception: {e}")

