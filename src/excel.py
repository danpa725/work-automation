from openpyxl import load_workbook


def get_column(dir, row_location):
    excel_data = load_workbook(dir, data_only=True)["Sheet1"]

    return list(
        filter(
            lambda x: x != None and "-" in x,
            map(lambda x: x.value, excel_data[row_location]),
        )
    )
