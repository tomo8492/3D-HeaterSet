from nicegui import ui

class CustomEditor(ui.editor):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.current_id = None