from guitest import main
import sys, database_op as _DO, signal
from sys_control import sigleton
from sys_control import get_frozen_stat

TERMINATE_SINGAL = 0

def signal_handler(sig, frame):
    global TERMINATE_SINGAL
    TERMINATE_SINGAL += 1
    if TERMINATE_SINGAL >= 10:
        sys.exit(0)

if __name__ in {"__main__", "__mp_main__"}:
    if get_frozen_stat():
        signal.signal(signal.SIGINT, signal_handler)

    command = sys.argv[1:]
    for com in command:
        if com == '--buildDB' or com == '-b':
            try:
                import pyi_splash
                pyi_splash.close()
            except:
                pass
            import tkinter
            verdict, result = _DO.create_table(True)
            from tkinter import messagebox
            rt = tkinter.Tk()
            rt.withdraw()
            message = 'データベース初期化処理結果：\n'
            if verdict:
                message += 'データベース初期化処理完了しました、この画面を閉じてください。'
            else:
                message += f'データベース初期化処理失敗しました、エラーメッセージ：{result}\nこの画面を閉じてください。'
            messagebox.showwarning('プロセスマネージャー', message)
            rt.destroy()
            sys.exit()

    sigleton_check = sigleton()

    if sigleton_check == 0:
        main()

    else:

        sys.exit()