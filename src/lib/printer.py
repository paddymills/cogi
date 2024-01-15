
import csv

from sys import stdout
from tabulate import tabulate


def print_to_source(results, as_csv=False, sumcols=None):
    if as_csv:
        csv.writer(stdout).writerows(results)
    else:
        header = results.pop(0)

        if sumcols:
            totals = [None for _ in header]
            for i, c in enumerate(header):
                # can index a sum column by index or header title
                if i in sumcols or c in sumcols:
                    totals[i] = sum([r[i] for r in results])

            results.append("\001")
            results.append(totals)

        print(tabulate(results, headers=header))
