from nicegui import ui

class CustomTabPanel(ui.tab_panel):
    def __init__(self, name):
        super().__init__(name)
        self.aggrid_instance: ui.aggrid = None
        self.aggrid_container: ui.element = None