import win32event, win32api, pywintypes
import os, sys

def set_up_env():
    storage = '.storage'

    if getattr(sys, 'frozen', False):
        thisfile = os.path.dirname(sys.executable)
        storage = os.path.join(thisfile, 'cache')
        
    os.environ['NICEGUI_STORAGE_PATH'] = storage

def get_frozen_stat() -> bool:
    return getattr(sys, 'frozen', False)




def sigleton() -> int:
    '''
    # Return Value
    1 = Instance already existing, quit current instance

    2 = Error creating mutex, quit current instance

    0 = Successfully created mutex and proceed to main program
    '''
    mutex_name = 'ODB2_MUTEX'
    try:
        mutex = win32event.CreateMutex(None, False, mutex_name)
        if win32api.GetLastError() == 183:
            print('プロセスマネージャーすでに実行中です、新規始めたプロセスを終了します。')
            #import tkinter
            #from tkinter import messagebox
            #rt = tkinter.Tk()
            #rt.withdraw()
            #messagebox.showwarning('プロセスマネージャー', 'プロセスマネージャーすでに実行中です、新規始めたプロセスを終了します。')
            try:
                #session = app.storage.general.get(USER_INS.uid)
                #if not session:
                import webbrowser
                webbrowser.open('http://localhost:8080/')
            except Exception as e:
                import tkinter
                from tkinter import messagebox
                rt = tkinter.Tk()
                rt.withdraw()
                messagebox.showwarning('プロセスマネージャー', f'実行中のプロセスマネージャーへのアクセスをトライしましたが、以下の原因で失敗しました。\nこの画面を閉じて再度プログラム立ち上げを試してください、エラーが繰り返される場合はIT管理者まで連絡ください。\n{e}')
                rt.destroy()
            return 1

        else:
            print('プロセスマネージャーを起動します。')
            return 0
    except pywintypes.error as e:
        print(f'Error creating mutex: {e}')
        return 2