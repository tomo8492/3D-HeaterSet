from nicegui import ui
from typing import Dict, Literal

class CustomTable(ui.table):
    def __init__(self, *, rows, columns = None, column_defaults = None, row_key = 'id', title = None, selection = None, pagination = None, on_select = None, on_pagination_change = None):
        super().__init__(rows=rows, columns=columns, column_defaults=column_defaults, row_key=row_key, title=title, selection=selection, pagination=pagination, on_select=on_select, on_pagination_change=on_pagination_change)
        self.indexed_column_name_dict : Dict[int, str] = None
        self.context_menu: Dict[Literal['copy'] | str, ui.menu_item] = {}
