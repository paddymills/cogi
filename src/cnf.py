
from argparse import ArgumentParser
from collections import defaultdict
from tqdm import tqdm

from lib.db import SndbConnection
from lib.parsers import SheetParser

def main():
    parser = ArgumentParser()
    parser.add_argument('--balance', help="Check SAP/Sigmanest confirmation balance")
    parser.add_argument('--raw', help="Check parts from raw query are confirmed")
    args = parser.parse_args()

    if args.balance:
        check_balance()
    else:
        check_raw_cnf()

def check_balance():
    data = SheetParser().parse_sheet(with_progress=True)
    parts = defaultdict(int)
    for row in data:
        if row.type == 'PP01':
            parts[row.part] += row.qty

    with SndbConnection() as db:
        c = db.cursor
        for part, qty in tqdm(parts.items(), desc="comparing SAP & Sigmanest"):
            c.execute("""
                SELECT QtyInProcess
                FROM PIPArchive
                WHERE PartName=? AND TransType='SN102'
            """, part.replace('-', '_', 1))
            try:
                cqty = c.fetchone()[0]
                # if qty < cqty:
                if qty != cqty:
                    tqdm.write("{} {:0.0f}/{}".format(part, qty, cqty))
                # else: print(part, "\tfull Qty confirmed")
            except TypeError:
                pass


def check_raw_cnf():
    parts = open('temp/parts.txt').read().split()

    data = SheetParser().parse_sheet(with_progress=True)
    for row in data:
        while row.part in parts:
            parts.remove(row.part)

    for part in parts:
        print(part)


if __name__ == "__main__":
    # main()
    check_balance()
