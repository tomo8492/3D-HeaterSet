class NoKeyExc(Exception):
    def __init__(self):
        super().__init__("No Valid Key.")

class ExtraKeyExc(Exception):
    def __init__(self):
        super().__init__("More than one key found.")

class FrequentExc(Exception):
    def __init__(self):
        super().__init__("Frequent Request.")

class NoDataExc(Exception):
    def __init__(self):
        super().__init__("Unable to retrive data.")

class InvalidInputExc(Exception):
    def __init__(self):
        super().__init__("Invalid input.")

class NoConnectionExc(Exception):
    def __init__(self):
        super().__init__("No internet connection.")

class NoDBConnectionExc(Exception):
    def __init__(self, error_mes="Error happening in SQL plain."):
        super().__init__(f"Unable to connect to DB, check the status of DB.\nDetail of exception:{error_mes}")

class NoItemSelectedExc(Exception):
    def __init__(self):
        super().__init__("No item selected.")

class ModuleOutlookFail(Exception):
    def __init__(self, error_mes=None):
        super().__init__(f"Something is wrong with Outlook Client, here is the details:{error_mes}")

class InternalErrorExc(Exception):
    def __init__(self, reason: str = ''):
        super().__init__(f"Internal error happened: {reason}")

class SQL_ErrorExc(Exception):
    def __init__(self, reason: str = ''):
        super().__init__(f"SQL error happened --> \n {reason} \n ##############")

class NoBookFound(Exception):
    def __init__(self, *args):
        super().__init__(*args)

class NoBookSet(Exception):
    def __init__(self, *args):
        super().__init__(*args)

