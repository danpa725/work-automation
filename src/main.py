import pygetwindow as gw
import pyautogui as pg
import time

from excel import get_column

BUTTON_POS = {
    'x': 174,
    'y': 204
}

WINDOW_TITLE = "노인복지종합시스템PLUS"


def automation(auto_list):
    focus_window = pg.getWindowsWithTitle(WINDOW_TITLE)[0]
    focus_window.maximize()

    if focus_window.isActive == False:
        focus_window.activate()

    time.sleep(1)

    for list in auto_list:
        pg.moveTo(focus_window.left +
                  BUTTON_POS['x'], focus_window.top + BUTTON_POS['y'])
        pg.click(focus_window.top +
                 BUTTON_POS['x'], focus_window.top + BUTTON_POS['y'])
        pg.write(list)
        pg.press('enter')
        time.sleep(1.5)


def iter(func, arg, iter):
    for _ in range(iter):
        func(arg)
    return


data = get_column("", "E")

iter(automation, data, 1)
