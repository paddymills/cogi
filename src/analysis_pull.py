
import datetime as dt
import os
import pyperclip
import xlwings as xl

from lib.db import SndbConnection
from lib.parsers import SheetParser

def main():
    wb = xl.Book(r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx")

    fill_sheet(wb)
    get_mb51_query_data(wb)


def fill_sheet(wb):
    sqlfile = os.path.join(os.path.dirname(__file__), "sql", "get_analysis_data.sql")

    sheet_name = monday()
    if sheet_name not in wb.sheet_names:
        wb.sheets['template'].copy(before=wb.sheets['Issues'], name=sheet_name)

    if wb.sheets[sheet_name].range("A2").value is None:
        data = SndbConnection().query_from_sql_file(sqlfile)
        wb.sheets[sheet_name].range("A2").value = [list(row) for row in data]
    else:
        print("\033[91m Sheet {} already exists and has data\033[00m".format(sheet_name))


def get_mb51_query_data(wb):
    sheet = wb.sheets[monday()]
    mm = []
    mm.extend(sheet.range("C2").expand('down').value)
    mm.extend(sheet.range("H2").expand('down').value)

    pyperclip.copy('\r\n'.join(sorted(set(mm))))
    print("Parts and Materials copied to clipboard")

def monday():
    today = dt.date.today()
    monday = today - dt.timedelta(days=today.weekday())

    return monday.strftime("%Y-%m-%d")


if __name__ == "__main__":
    main()
