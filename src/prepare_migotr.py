
import xlwings
from lib.parsers import SheetParser
from collections import defaultdict

parser = SheetParser()
wb = xlwings.Book('cogi_mb52_compare.xlsx')

# MB52: current stock
_stock = list( SheetParser(wb.sheets['mb52']).parse_sheet(with_progress=True) )
stock = defaultdict(lambda: defaultdict(int))
for s in _stock:
    stock[s.part][s.wbs] += s.qty

# COGI: error items
data = iter(SheetParser(wb.sheets['cogi']).parse_sheet(with_progress=True))
cogi = [ next(data) ]
for c in data:
    if cogi[-1].part == c.part and cogi[-1].wbs == c.wbs and cogi[-1].plant == c.plant:
        cogi[-1].qty += c.qty
    else:
        cogi.append(c)

# remove items not in MB52
cogi = [c for c in cogi if c.part in stock]

def reduction(c, wbs):
    reduction_qty = min( stock[c.part][wbs], c.qty )
    stock[c.part][wbs] -= reduction_qty
    c.qty -= reduction_qty

    if stock[c.part][wbs] == 0:
        del stock[c.part][wbs]

    return reduction_qty

def print_stock():
    for s in stock:
        print(s)
        for k, v in stock[s].items():
            print('\t', k, v)

migotr = list()
print_stock()
# TODO: resolve deadlock situation with demand
#
# | part | MB52 | Needs |
# |------|------|-------|
# |  x1a |  3   |   4   |
# |  x1a |  1   |   2   |

# reduce quantity that already satisfies demand
for c in cogi:
    if c.wbs in stock[c.part]:
        reduction(c, c.wbs)

print('===========')
print_stock()
for c in sorted(cogi, key=lambda x: (x.part, x.qty)):
    # print(c)
    while c.qty > 0 and len(stock[c.part]) > 0:
        key = next(iter(stock[c.part].keys()))
        qty = reduction(c, key)

        migotr.append([c.part, None, c.plant, 'PROD', qty, key, None, 'PROD', c.wbs])
        if len(migotr) % 23 == 0:
            migotr.append([None] * 9)

# wb.sheets['migotr'].range('A:I').clear_contents()
wb.sheets['migotr'].range('A2').value = migotr
