
from argparse import ArgumentParser
from collections import defaultdict
from datetime import datetime
from pprint import pprint
from tabulate import tabulate, SEPARATING_LINE
# import xlwings

from lib.db import SndbConnection
from lib.parsers import SheetParser

cutoff = datetime(year=2023, month=1, day=1)
reportable_variance = 1

query = """
DECLARE @LastIntervalHourTimestamp DATETIME
DECLARE @intervalHours INT
SELECT @intervalHours = 4
SELECT @LastIntervalHourTimestamp = DATEADD(HOUR, DATEDIFF(HOUR, 0, GETDATE()) / @intervalHours * @intervalHours, 0)

SELECT
    REPLACE(PartName, '_', '-') AS Part,
    part.ProgramName AS Program,
    QtyProgram,
    NestedArea,
    NestedArea * QtyProgram AS TotalNestedArea
FROM PartArchive AS part
	inner join StockArchive as stock
		on part.ArchivePacketID=stock.ArchivePacketID
	inner join ProgArchive as program
		on part.ArchivePacketID=program.ArchivePacketID
WHERE stock.primecode = ?
AND part.ArcDateTime >= ?
AND part.ArcDateTime < @LastIntervalHourTimestamp
AND program.TransType = 'SN102'
"""


def main():
    parser = ArgumentParser()
    parser.add_argument("--orders", action="store_true", help="Generate list of order")
    args = parser.parse_args()

    if args.orders:
        get_orders()
    
    else:
        compare()


def compare():
    print = pprint

    orders = dict()
    for row in parse_cohv():
        orders[row.order] = row.part

    mm = None
    parts = defaultdict(float)
    issued = defaultdict(float)
    for row in parse_mb51():
        mm = row.material

        # row.qty is negative for a consumption
        match row.type:
            case '201' | '202':
                # issue to cost center
                issued[row.program] += row.qty
            case '221' | '222':
                # issue to project
                issued[row.program] += row.qty
            case '261' | '262':
                # issue to order
                parts[orders[row.order]] += row.qty
            case x:
                print(f"Unmatched {x} ({type(x)})")

    # since all consumptions are negative,
    #  we can counter this by using addition again
    variance = [["Program", "Part", "Qty"]]
    for row in get_sn_data(mm):
        if row.Part in parts:
            parts[row.Part] += row.TotalNestedArea
        elif row.Program in issued:
            issued[row.Program] += row.TotalNestedArea
        else:
            variance.append((row.Program, row.Part, row.TotalNestedArea))

    for part, qty in parts.items():
        if abs(qty) > reportable_variance:
            variance.append(('', part, qty))
    for program, qty in issued.items():
        if abs(qty) > reportable_variance:
            variance.append((program, '', qty))

    with open('temp/results.txt', 'w') as f:
        total = sum([x[-1] for x in variance[1:]])
        table = [*variance, SEPARATING_LINE, ("Total", "", total)]
        f.write(tabulate(table, headers="firstrow"))
    with open('temp/results.csv', 'w') as f:
        it = iter(variance)
        f.write("{},{},{}\n".format(*next(it)))
        for line in it:
            f.write("{},{},{:.3f}\n".format(*line))



def get_orders():
    with open('temp/orders.txt', 'w') as f:
        orders = set([row.order for row in parse_mb51() if row.order])
        f.write('\n'.join(sorted(orders)))


def parse_mb51():
    wb = SheetParser(wb='mb51.xlsx')

    for row in wb.parse_sheet(with_progress=True):
        if row.date >= cutoff:
            yield row


def parse_cohv():
    return SheetParser(wb='cohv.xlsx').parse_sheet(with_progress=True)


def get_sn_data(mm):
    with SndbConnection() as db:
        db.cursor.execute(query, mm, cutoff.strftime('%Y-%m-%d'))

        for row in db.cursor.fetchall():
            yield row


if __name__ == "__main__":
    main()
