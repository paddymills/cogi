
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


def sap_sn_compare():
    program_pattern = re.compile(r"\d{5}")

    orders = parse_cohv()
    cons: dict[str, Mb51Item] = dict()
    for x in parse_mb51(orders):
        if x.name in cons:
            cons[x.name].qty += x.qty
            cons[x.name].area += x.area
        else:
            cons[x.name] = x

    mm = SheetParser(wb='mb51.xlsx').parse_row().matl

    for r in get_sn_data(mm):
        if r.Part in cons:
            cons[r.Part].qty -= r.Qty
            cons[r.Part].area += r.TotalNestedArea
        elif r.Program in cons:
            cons[r.Program].qty -= 1
            cons[r.Program].area += r.TotalNestedArea
        else:
            cons[r.Part] = Mb51Item(r.Part, r.Qty, r.TotalNestedArea)

    table = [["Part/Program", "Qty", "Area"]]
    for v in cons.values():
        if abs(v.area) >= reportable_variance and \
            v.qty > 0:
            table.append([v.name, v.qty, v.area])

    print(table_with_totals(table, totals_index=[1, 2]))


def sap_file_compare():
    program_pattern = re.compile(r"\d{5}")

    # mm = SheetParser(wb='mb51.xlsx').parse_row().matl
    mm = '50/50W-0108'

    orders = parse_cohv()
    cons = list(parse_mb51(mm, orders))

    matched = list()
    not_matched = list()

    for row in tqdm(parse_issue(), desc="removing successful Issue items"):
        if row.matl == mm:
            for i, c in enumerate(cons):
                if c.name == row.prog and c.area + row.qty == 0:
                    matched.append(row)
                    cons.pop(i)
                    break
            else:
                not_matched.append(row)
    
    for row in tqdm(parse_production(), desc="removing successful Production items"):
        if row.matl == mm:
            for i, c in enumerate(cons):
                if c.name == row.part and c.qty == row.pqty and c.area + row.mqty == 0:
                    matched.append(row)
                    cons.pop(i)
                    break
            else:
                not_matched.append(row)

    # rollup consumption
    temp = dict()
    for c in cons:
        if c.name not in temp:
            temp[c.name] = c
        else:
            temp[c.name].qty += c.qty
            temp[c.name].area += c.area
    cons = list(temp.values())

    # rollup not_matched
    nm = list()
    for row in tqdm(not_matched, desc="removing successful Production items (rollup)"):
        match row:
            case ProductionItem(part, prog, _, _, _):
                names = (part, prog)
            case IssueItem(_, prog, _):
                names = (prog)
        for i, c in enumerate(cons):
            if c.name in names: # \
                # and c.qty + row.pqty >= 0 \
                # and c.area + row.mqty >= -1 * reportable_variance:

                cons[i].qty -= row.pqty
                cons[i].area += row.mqty
                matched.append(row)
                break
        else:
            nm.append(row)
    

    write_table(cons, ["Part/Program", "Qty", "Area"], 'temp/cons.txt')
    write_table(matched, ["Part", "Program", "Qty", "Area"], 'temp/matched.txt')
    write_table(nm, ["Part", "Program", "Qty", "Area"], 'temp/nm.txt')


def parse_production():
    progs = list()

    path = os.path.join(os.environ['USERPROFILE'], r"Documents\sapcnf\outbound\Production_*.outbound.archive")
    for fn in glob(path):
        with open(fn) as f:
            file_progs = list()
            for line in f.readlines():
                s = line.strip().split('\t')
                if len(s) > 10:
                    part = s[0]
                    pqty = float(s[4])
                    matl = s[6]
                    mqty = float(s[8])
                    prog = s[12]

                    if prog in progs:
                        continue

                    # if part == '1200037B-X201A':
                    #     tqdm.write(f"found {part} in {fn} for program {prog} ({prog in progs})")

                    file_progs.append(prog)
                    yield ProductionItem(part, prog, matl, pqty, mqty)

        progs.extend(file_progs)

def parse_issue():
    progs = list()

    path = os.path.join(os.environ['USERPROFILE'], r"Documents\sapcnf\outbound\Issue_*.archive_*")
    for fn in glob(path):
        with open(fn) as f:
            file_progs = list()
            for line in f.readlines():
                s = line.strip().split('\t')
                if len(s) > 9:
                    matl = s[3]
                    qty  = float(s[5])
                    prog = s[9]

                    if prog in progs:
                        continue

                    file_progs.append(prog)
                    yield IssueItem(matl, prog, qty)

        progs.extend(file_progs)


def parse_mb51(mm: str, orders: dict) -> dict[str, Mb51Item]:

    wb = SheetParser(wb='mb51.xlsx')
    for row in wb.parse_sheet(with_progress=True):
        if row.matl != mm:
            continue
        if row.date >= cutoff:
            # row.qty is negative for a consumption
            area = row.qty
                
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


def table_with_totals(data, totals_index=[-1]):
    table = [*data, SEPARATING_LINE]

    total_rows = [
        ("Sn side", lambda x: x > 0),
        ("SAP side", lambda x: x < 0),
        ("Total", lambda _: True),
    ]

    for title, fn in total_rows:
        row = [title, *[None] * (len(data[0])-1)]

        for i in totals_index:
            row[i] = sum([x[i] for x in data[1:] if fn(x[i])])

        if any(row[1:]) > 0:
            table.append(row)
    
    return tabulate(table, headers="firstrow")


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

    if len(table) > 1:
        totals_index = [i for i, x in enumerate(table[1]) if type(x) in (int, float)]
    else:
        totals_index = []

    with open(filename, 'w') as f:
        f.write(table_with_totals(table, totals_index))


if __name__ == "__main__":
    # sap_sn_compare()
    sap_file_compare()
