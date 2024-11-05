import datetime as dt
import os
import sys
import pyperclip
import xlwings as xl

from lib.db import SndbConnection
from lib.parsers import SheetParser


def main():
    wb = xl.Book(
        r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx"
    )

    fill_sheet(wb)


def fill_sheet(wb):
    sqlfile = os.path.join(os.path.dirname(__file__), "sql", "get_analysis_data.sql")

    sheet_name = monday()
    if sheet_name not in wb.sheet_names:
        wb.sheets["template"].copy(before=wb.sheets["Issues"], name=sheet_name)

    if wb.sheets[sheet_name].range("A2").value is None:
        data = SndbConnection().query_from_sql_file(sqlfile)
        wb.sheets[sheet_name].range("A2").value = [list(row) for row in data]

        get_mb51_query_data(wb)
    else:
        print(
            "\033[91m Sheet {} already exists and has data\033[00m".format(sheet_name)
        )
        get_not_filled(wb)


def get_mb51_query_data(wb):
    sheet = wb.sheets[monday()]
    mm = []
    mm.extend(sheet.range("C2").expand("down").value)
    mm.extend(sheet.range("H2").expand("down").value)
    earliest_data = min(sheet.range("B2").expand("down").value)

    pyperclip.copy("\r\n".join(sorted(set(mm))))
    print(
        "Parts and Materials copied to clipboard. Earliest date is {}".format(
            earliest_data.strftime("%m-%d-%Y")
        )
    )


def get_not_filled(wb):
    sheet = wb.sheets[monday()]
    mm = list()
    for r in sheet.range("A2:L2").expand("down").value:
        if r[-1] in (None, ""):
            mm.append(r[2])
            mm.append(r[7])

    dates = sheet.range("B2").expand("down").value

    pyperclip.copy("\r\n".join(sorted(set(mm))))
    print(
        "Not matched Parts and Materials copied to clipboard. Date range is {} to {}".format(
            min(dates).strftime("%m-%d-%Y"), max(dates).strftime("%m-%d-%Y")
        )
    )


def monday():
    if len(sys.argv) > 1:
        return sys.argv[1]

    today = dt.date.today()
    monday = today - dt.timedelta(days=today.weekday())

    return monday.strftime("%Y-%m-%d")


if __name__ == "__main__":
    main()
