from tkinter import messagebox
from openpyxl import load_workbook


def get_column(directory: str, row_location: str):
    try:
        excel_data = load_workbook(directory, data_only=True)["Sheet1"]
        column_data = list(
            filter(
                lambda x: x != None and "-" in x,
                map(lambda x: x.value, excel_data[row_location]),
            )
        )

        return column_data
    except:
        messagebox.showerror(
            "error", "액셀 파일 데이터가 유효한지 확인하세요, 아이디는 XXXX-XXXXX 형식의 순수 문자열이어야 합니다."
        )
        return
