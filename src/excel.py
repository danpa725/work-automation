from openpyxl import load_workbook


def get_column(directory: str, row_location: str):
    excel_data = load_workbook(directory, data_only=True)["Sheet1"]
    column_data = list(
        filter(
            lambda x: x != None and "-" in x,
            map(lambda x: x.value, excel_data[row_location]),
        )
    )

    return column_data
