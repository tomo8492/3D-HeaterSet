from nicegui import ui

class CustomDialog(ui.dialog):
    def __init__(self, *, value = False):
        super().__init__(value=value)
        self.inited_tab = {'作業実績' : None,
                           '作業予定' : None,
                           '出荷制限' : None,
                           'シリアル' : None}