from nicegui import ui

class CustomExpansion(ui.expansion):
    def __init__(self, text = '', *, caption = None, icon = None, group = None, value = False, on_value_change = None):
        super().__init__(text, caption=caption, icon=icon, group=group, value=value, on_value_change=on_value_change)
        self.value_collection = None
        self.value_collection_2 = None