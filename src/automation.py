from typing import Dict
import pyautogui as pg
from tkinter import messagebox
from time import sleep
from excel import get_column

START_BUTTON_POS = {"x": 79, "y": 156}
MAX_WINDOW_SIZE_POS = {"x": 820, "y": 65}
INPUT_BUTTON_POS = {"x": 160, "y": 200}

WINDOW_TITLE = "노인복지종합시스템PLUS"


def move_click(window_object: any, pos: Dict[str, str]):
    pg.moveTo(window_object.left + pos["x"], window_object.top + pos["y"])
    pg.click(window_object.top + pos["x"], window_object.top + pos["y"])


def move_dbclick(window_object: any, pos: Dict[str, str]):
    pg.moveTo(window_object.left + pos["x"], window_object.top + pos["y"])
    pg.doubleClick(window_object.top + pos["x"], window_object.top + pos["y"])


def automation(exceldir: str, id_location: str, iter_time: int):
    try:
        focus_window = pg.getWindowsWithTitle(WINDOW_TITLE)[0]
        focus_window.maximize()

        if focus_window.isActive == False:
            focus_window.activate()

        # extract id list
        id_list = get_column(exceldir, id_location)
        if len(id_list) == 0:
            messagebox.showerror("error", "해당 열에 데이터가 없습니다.")
            return

        move_click(focus_window, START_BUTTON_POS)
        sleep(3.5)

        move_dbclick(focus_window, MAX_WINDOW_SIZE_POS)
        sleep(2.5)

        for id in id_list:
            move_click(focus_window, INPUT_BUTTON_POS)

            pg.press("right", presses=10)
            pg.press("backspace", presses=11)

            pg.write(id)

            pg.press("enter")

            sleep(0.5)

            pg.press("esc")

            sleep(iter_time)

    except:
        messagebox.showerror("error", "회원 회비입금관리가 최대로 열려 있는지 확인하세요")
        return
