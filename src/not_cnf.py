import xlwings
from os import path

from lib import parsers
from lib.db import SndbConnection

from re import compile as regex
from collections import defaultdict
from itertools import groupby
from types import SimpleNamespace
from datetime import datetime
from tabulate import tabulate
from tqdm import tqdm


def main():
    # regexes
    CNF = regex(r"\d{10}")
    NOT_CNF = regex(r"\d{6,7}")

    # parse active xl sheet
    cohv = parsers.SheetParser().parse_sheet(with_progress=True)

    # parse out dict structure
    cnf = defaultdict(int)
    not_cnf = defaultdict(list)
    for x in tqdm(list(cohv), desc="Reading orders"):
        if x.order is None: # if an empty row was parsed
            break

        if x.type == 'PP01':
            cnf[x.matl] += x.qty
        elif x.type == 'PR':
            not_cnf[x.matl].append((x.wbs, x.qty, x.plant))
        else:
            print("Order not matched:", x.order)

    # add keys to cnf that exist in not_cnf
    # since cnf is a defaultdict, we just need to access the element to add it
    for key in not_cnf:
        cnf[key] += 0

    reader = SnReader()
    confirmations = list()
    for part, qty_confirmed in tqdm(cnf.items(), desc="Comparing COHV"):
        if part not in not_cnf:
            continue

        qty_burned = reader.get_part_burned_qty(part)
        # print("{} cnf: {} burned: {}".format(part, qty_confirmed, qty_burned))

        if qty_confirmed < qty_burned:
            qty_to_confirm = qty_burned - qty_confirmed
            for wbs, qty, plant in sorted(not_cnf[part]):
                if qty < qty_to_confirm:
                    confirmations.append((part, wbs, qty, plant))
                    qty_to_confirm -= qty
                else:  # open quantity is equal or greater than what needs to be confirmed
                    confirmations.append((part, wbs, qty_to_confirm, plant))
                    break

        try:
            del not_cnf[part]
        except KeyError:
            pass

    # print(tabulate([k, sum([x[2] for x in v])] for k,v in groupby(confirmations, lambda x: x[0])))

    processed_lines = parsers.CnfFileParser().get_cnf_file_rows([x[0] for x in confirmations])

    index = SimpleNamespace(matl=0, qty=4, wbs=2, plant=11, in2_consumption=8)
    # matl = SimpleNamespace(matl=6, qty=8, loc=10, wbs=7, plant=11)

    templates = dict()
    for line in processed_lines:
        part = line[index.matl]

        templates[part] = line

    # create output file
    ready_file = path.join(
        path.dirname(path.realpath(__file__)),
        "Production_{:%Y%m%d%H%M%S}.ready".format(datetime.now())
    )

    if not confirmations:
        print("Nothing to confirm")
        return

    not_in_templates = list()
    cnf_data = list()

    for part, wbs, qty, plant in confirmations:
        if part not in templates:
            not_in_templates.append(part)
            continue

        line = templates[part]
        in2_per_ea = float(line[index.in2_consumption]) / int(line[index.qty])

        line[index.wbs] = wbs
        line[index.qty] = str(int(qty))
        line[index.in2_consumption] = str(round(in2_per_ea * int(qty), 3))
        line[index.plant] = plant

        cnf_data.append("\t".join(line))

    if cnf_data:
        with open(ready_file, 'w') as cnf_file:
            for line in cnf_data:
                cnf_file.write(line)

    if not_in_templates:
        print("Parts not found in cnf files:")
        print("\t" + "\n\t".join(not_in_templates))


class SnReader(SndbConnection):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def get_part_burned_qty(self, part_name):
        # blacklist_workorders = ('REMAKES', 'EXTRAS')
        blacklist_workorders = ('EXTRAS')

        # part_query = part_name.replace("-", "_", 1)
        part_query = '%' + part_name[4:].replace("-", "%", 1)

        sql_file = path.join(path.dirname(__file__), "sql", "part_burned_qty.sql")
        self.execute_sql_file(sql_file, part_query, self.simtrans_cutoff)

        total = 0
        for wo, qty in self.cursor.fetchall():
            if wo not in blacklist_workorders:
                total += qty

        return total

    @property
    def simtrans_cutoff(self):
        # sim trans files are generated every 4 hours
        # 00:00, 04:00, 08:00, 12:00, 16:00, 20:00
        # (generated on the half hour, but data is up to the hour)

        now = datetime.now()

        last_run = now.replace(hour=now.hour - now.hour % 4)

        return "{0:%m}/{0:%d}/{0:%Y} {0:%I}:00:00 {0:%p}".format(last_run)


if __name__ == "__main__":
    main()
