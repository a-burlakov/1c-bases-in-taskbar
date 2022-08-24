import win32gui
import win32process
from psutil import Process
import schedule
import time


def add_ibname_to_1C_window(hwnd, more) -> None:
    '''
    Если окно является окном 1С,
    то добавляет слева к его названию имя открытой базы.
    :param hwnd: Открытое окно Windows
    '''

    # Берем только окна, связанные с 1С.
    class_name = win32gui.GetClassName(hwnd)
    if 'V8TopLevelFrame' not in class_name:
        return

    window_text = win32gui.GetWindowText(hwnd)

    # Находим процесс, и по этому процессу ищем командную строку.
    process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
    process = Process(process_id)
    commandline = process.cmdline()

    # Ищем имя базы.
    iterator = (x for x in commandline if x.startswith('/IBName'))
    ib_name = next(iterator, '!None')
    if ib_name != '!None':
        ib_name = ib_name.replace('/IBName', '')
    else:
        return

    # Устанавливаем новое имя.
    if ib_name not in window_text:
        win32gui.SetWindowText(hwnd, f'{ib_name} - {window_text}')


def windows_passage() -> None:
    '''
    Проходит по открытым окнам с указанной функцией.
    '''
    win32gui.EnumWindows(add_ibname_to_1C_window, None)


if __name__ == '__main__':
    schedule.every(2).seconds.do(windows_passage)

    while True:
        schedule.run_pending()
        time.sleep(2)