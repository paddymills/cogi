
import xlwings

from lib.db import SndbConnection

with SndbConnection() as db:
    db.cursor.execute("SELECT * FROM Stock")
    stock = db.collect_table_data()

    db.cursor.execute("SELECT * FROM Remnant")
    remnant = db.collect_table_data()

    wb = xlwings.Book()
    wb.sheets[0].name = "Stock"
    wb.sheets[0].range("A1").value = stock

    wb.sheets.add(after="Stock")
    wb.sheets[1].name = "Remnant"
    wb.sheets[1].range("A1").value = remnant

