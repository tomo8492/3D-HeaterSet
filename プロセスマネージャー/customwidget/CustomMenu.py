from nicegui import ui

class CustomMenu(ui.menu):
    def __init__(self, *, value = False):
        super().__init__(value=value)
        self.data = None
        self.skeleton = None
