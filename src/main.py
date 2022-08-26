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
    ib_name = ib_name_in_commandline(commandline)
    if ib_name is None:
        return

    # Устанавливаем новое имя.
    if not window_text.startswith(ib_name):
        win32gui.SetWindowText(hwnd, f'{ib_name} - {window_text}')


def ib_name_in_commandline(commandline: list) -> str | None:
    '''
    Ищет имя информационной базы из командной строки 1С.
    :param commandline: Список строк, из которых состоит строка параметров процесса.
    :return:
    '''

    # Поиск по параметру "/IBName".
    # Вид строки: "/IBNameMYBASE".
    iter = (x for x in commandline if x.startswith('/IBName'))
    par = next(iter, '')
    if par:
        ib_name = par.replace('/IBName', '')
        ib_name = ib_name.strip(' "\'')
        return ib_name

    # Поиск по параметру "/S".
    # Вид строки: '/S"server-1C:1641\ZUP"'
    iter = (x for x in commandline if x.startswith('/S') and '\\' in x)
    par = next(iter, '')
    if par:
        ib_name = par.split('\\')[-1]
        ib_name = ib_name.strip(' "\'')
        return ib_name

    # Поиск по параметру "/F"
    # Вид строки: '/F"D:\1C_base\ZUPRAZR"'
    iter = (x for x in commandline if x.startswith('/F') and '\\' in x)
    par = next(iter, '')
    if par:
        ib_name = par.split('\\')[-1]
        ib_name = ib_name.strip(' "\'')
        return ib_name

    # Поиск по параметру "/IBConnectionString".
    # Вид строки: "Srvr=MYSERVER;Ref=MYBASE;".
    iter = (x.lower() for x in commandline if 'ref=' in x.lower())
    par = next(iter, '')
    if par:
        par = par.partition(';ref=')[2] # убираем всё до ";ref="
        par = par.partition(';')[0] # убираем всё после ";"
        ib_name = par.strip(' "\'')
        return ib_name


def windows_passage() -> None:
    '''
    Проходит по открытым окнам с указанной функцией.
    '''
    win32gui.EnumWindows(add_ibname_to_1C_window, None)


if __name__ == '__main__':

    istest = False
    if istest:
        windows_passage()
        raise SystemExit('Тестовый прогон закончен.')

    schedule.every(2).seconds.do(windows_passage)

    while True:
        schedule.run_pending()
        time.sleep(2)
