
from argparse import ArgumentParser
from collections import defaultdict
from datetime import datetime
from pprint import pprint
from tabulate import tabulate, SEPARATING_LINE
from tqdm import tqdm
# import xlwings

import os

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
    parser.add_argument("--single", action="store_true", help="Run compare for a single mm")
    parser.add_argument("--consumption", action="store_true", help="Check consumption")
    parser.add_argument("--confirmation", action="store_true", help="Check confirmations")
    parser.add_argument("--mm", action="store_true", help="generate list of material masters")
    args = parser.parse_args()

    if args.orders:
        get_orders()
    
    elif args.mm:
        get_mms()
    
    elif args.single and args.consumption:
        compare_single()

    elif args.consumption:
        compare_many()

    elif args.confirmation:
        check_confirmations()


def compare_many(mm=None, show_unmatched=False):
    # print = pprint
    for f in os.scandir("temp/per_mm"):
        os.remove(f.path)

    orders = dict()
    for row in parse_cohv():
        orders[row.order] = row.part

    mm = None
    table = [["Material", "Qty"]]
    partnames = list()
    for row in parse_mb51():
        if row.material != mm:
            # add new line to report
            if mm is not None:
                variance = calc_variance(mm, parts, issued)
                total = sum([x[-1] for x in variance[1:]])
                if total > reportable_variance:
                    table.append([mm, total])

                    for _, part, _ in variance[1:]:
                        partnames.append(part)

                    with open('temp/manyresults.txt', 'w') as f:
                        f.write(tabulate(table, headers="firstrow"))
                    with open(f'temp/per_mm/{mm.replace("/", "_")}.txt', 'w') as f:
                        f.write(table_with_totals(variance))

            # reset collections
            parts = defaultdict(float)
            issued = defaultdict(float)

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
                try:
                    parts[orders[row.order]] += row.qty
                except KeyError:
                    pass
            case x if show_unmatched:
                print(f"Unmatched {x} ({type(x)})")

    print(tabulate(table, headers="firstrow"))
    with open('temp/parts.txt', 'w') as f:
        f.write("\n".join(sorted(set(partnames))))
    


def compare_single(mm=None, show_unmatched=False):
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
            case x if show_unmatched:
                print(f"Unmatched {x} ({type(x)})")

    variance = calc_variance(mm, parts, issued)
    with open('temp/results.txt', 'w') as f:
        # total = sum([x[-1] for x in variance[1:]])
        # table = [*variance, SEPARATING_LINE, ("Total", "", total)]
        f.write(table_with_totals(variance))
    with open('temp/results.csv', 'w') as f:
        it = iter(variance)
        f.write("{},{},{}\n".format(*next(it)))
        for line in it:
            f.write("{},{},{:.3f}\n".format(*line))


def check_confirmations():
    confirmations = defaultdict(int)
    planned = defaultdict(int)

    # pre-fill confirmations with parts from sigmanest
    with open("temp/parts.txt") as f:
        for row in f.readlines():
            confirmations[row.strip()] = 0

    for line in SheetParser(wb="cohv").parse_sheet():
        match line.type:
            case "PP01":    # Production Order
                confirmations[line.part] += line.qty
            case "PR":      # Planned Order
                planned[line.part] += line.qty

    underconsumption = list()
    table = [["Part", "Cnf", "Burned"]]
    with SndbConnection() as db:
        for k, v in tqdm(confirmations.items(), desc="checking confirmation balance"):
            db.execute("""
                       SELECT ISNULL(SUM(QtyProgram), 0) AS Qty
                       FROM PartArchive
                       WHERE PartName=? AND WoNumber != 'REMAKES'
                       """, k.replace("-", "_", 1))

            qty = db.fetchone().Qty
            if qty > v:
                if k in planned:
                    tqdm.write(f"{k} underconfirmed, but there are planned orders")
                underconsumption.append(k)
                table.append([k, v, qty])

    print(tabulate(table, headers="firstrow"))
    
    with open("temp/underconfirmation.txt", 'w') as f:
        f.write("\n".join(sorted(underconsumption)))


def calc_variance(mm, parts, issued):
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

    return variance


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


def table_with_totals(data, totals_index=[-1]):
    table = [*data, SEPARATING_LINE, ["Total", *[None] * (len(data[0])-1)]]
    for i in totals_index:
        total = sum([x[i] for x in data[1:]])
        table[-1][i] = total
    
    return tabulate(table, headers="firstrow")

def get_mms():
    print("not implemented!")


if __name__ == "__main__":
    main()
