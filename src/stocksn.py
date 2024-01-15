
from argparse import ArgumentParser
from collections import defaultdict
from datetime import datetime
from tabulate import tabulate, SEPARATING_LINE
from tqdm import tqdm

from dataclasses import dataclass
from glob import glob

import re
import os

from lib.db import SndbConnection
from lib.parsers import SheetParser

cutoff = datetime(year=2023, month=1, day=1)
reportable_variance = 10

query = """
DECLARE @LastIntervalHourTimestamp DATETIME
DECLARE @intervalHours INT
SELECT @intervalHours = 4
SELECT @LastIntervalHourTimestamp = DATEADD(HOUR, DATEDIFF(HOUR, 0, GETDATE()) / @intervalHours * @intervalHours, 0)

SELECT
    REPLACE(PartName, '_', '-') AS Part,
    part.ProgramName AS Program,
    QtyProgram AS Qty,
    NestedArea,
    NestedArea * QtyProgram AS TotalNestedArea
FROM PartArchive AS part
	inner join StockArchive as stock
		on part.ArchivePacketID=stock.ArchivePacketID
WHERE stock.primecode = ?
AND part.ArcDateTime >= ?
AND part.ArcDateTime < @LastIntervalHourTimestamp
"""

@dataclass
class Mb51Item:
    name: str
    qty: int
    area: float

@dataclass
class IssueItem:
    matl: str
    prog: str
    qty: float

@dataclass
class ProductionItem:
    part: str
    prog: str
    matl: str
    pqty: float
    mqty: float

@dataclass
class DbItem:
    part: str
    prog: str
    matl: str
    pqty: float
    mqty: float



def sap_sndb_compare():
    cons = dict()
    for row in parse_mb51(parse_cohv()):
        if row.name in cons:
            cons[row.name].qty += row.qty
            cons[row.name].area += row.area
        else:
            cons[row.name] = row

    mm = SheetParser(wb='mb51.xlsx').parse_row().matl
    not_matched = list()

    for row in get_sn_data(mm):
        if row.Part in cons:
            cons[row.Part].qty -= row.Qty
            cons[row.Part].area -= row.TotalNestedArea
        elif row.Program in cons:
            cons[row.Program].qty -= 1
            cons[row.Program].area -= row.TotalNestedArea

        else:
            not_matched.append(DbItem(row.Part, row.Program, mm, row.Qty, row.TotalNestedArea))
    

    nm = list()
    for x in not_matched:
        nm.append([x.part, x.prog, x.pqty, x.mqty])

    consp = list()
    for x in cons.values():
        if x.qty != 0 and abs(x.area) > reportable_variance:
            consp.append([x.name, x.qty, x.area])

    print(table_with_totals(consp, [1, 2], header=["Part/Program", "Qty", "Area"]))
    print(table_with_totals(nm, [2, 3], header=["Part", "Program", "Qty", "Area"]))


def parse_mb51(orders: dict) -> [Mb51Item]:

    wb = SheetParser(wb='mb51.xlsx')
    for row in wb.parse_sheet(with_progress=True):
        if row.date >= cutoff:
            # row.qty is negative for a consumption
            area = -1 * row.qty
                
            match row.type:
                case '201' | '202' | '221' | '222' if row.program is not None:
                    # issue to cost center
                    name = row.program
                    qty = 1
                case '261' | '262':
                    # issue to order
                    try:
                        name, qty = orders[row.order]
                    except KeyError:
                        tqdm.write(f"part not found for order: {row.order}")
                case _:
                    continue

            yield Mb51Item(name, qty, area)



def parse_cohv() -> dict[str, (str, int)]:
    orders = dict()
    for row in SheetParser(wb='cohv.xlsx').parse_sheet(with_progress=True):
        orders[row.order] = (row.part, row.qty)

    return orders


def get_sn_data(mm):
    with SndbConnection() as db:
        db.cursor.execute(query, mm, cutoff.strftime('%Y-%m-%d'))

        for row in db.cursor.fetchall():
            yield row


def table_with_totals(data, totals_index=[-1], header=[]):
    table = [*data, SEPARATING_LINE]

    total_rows = [
        ("Sn side", lambda x: x > 0),
        ("SAP side", lambda x: x < 0),
        ("Total", lambda _: True),
    ]

    for title, fn in total_rows:
        row = [title, *[None] * (len(header)-1)]

        for i in totals_index:
            row[i] = sum([x[i] for x in data if fn(x[i])])

        if any(row[1:]) > 0:
            table.append(row)
    
    return tabulate(table, headers=header)


def write_table(data, header, filename):
    def key(v):
        match v:
            case Mb51Item(name, _, _):
                return (name, None)
            case ProductionItem(part, prog, _, _, _):
                return (part, prog)
            case IssueItem(_, prog, _):
                return (prog, None)

    table = [header]
    for v in sorted(data, key=key):
        match v:
            case Mb51Item(name, qty, area):
                table.append([name, qty, area])
            case ProductionItem(part, prog, _, pqty, mqty):
                table.append([part, prog, pqty, mqty])
            case IssueItem(_, prog, qty):
                table.append([None, prog, 1, qty])

    totals_index = [i for i, x in enumerate(table[1]) if type(x) in (int, float)]

    with open(filename, 'w') as f:
        f.write(table_with_totals(table, totals_index))


if __name__ == "__main__":
    sap_sndb_compare()
