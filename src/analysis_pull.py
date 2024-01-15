
import datetime as dt
import pyperclip
import xlwings as xl

from lib.db import SndbConnection
from lib.parsers import SheetParser

def main():
    wb = xl.Book(r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx")

    if monday() not in wb.sheets:
        fill_sheet(wb)

    get_mb51_query_data(wb)


def fill_sheet(wb):
    sheet = wb.sheets['template'].copy(before='Issues', name=monday())
    sheet.range("A2").value = SndbConnection().query_from_sql_file(r'sql\get_analysis_data.sql')


def get_mb51_query_data(wb):
    sheet = wb.sheets[monday()]
    mm = []
    mm.extend(sheet.range("C2").expand('down').value)
    mm.extend(sheet.range("H2").expand('down').value)

    pyperclip.copy('\n'.join(mm))
    print("Parts and Materials copied to clipboard")

def monday():
    today = dt.date.today()
    monday = today - dt.timedelta(days=today.weekday())

    return monday.strftime("%Y-%m-%d")


if __name__ == "__main__":
    main()
