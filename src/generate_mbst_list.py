
from collections import defaultdict
from itertools import combinations
from tabulate import tabulate
from tqdm import tqdm
import xlwings

def gi():
    wb = xlwings.books["gi.xlsx"]
    data = wb.sheets.active.range("A1").expand().value
    header = data.pop(0)

    p = header.index("ProgramName")
    a = header.index("Total")

    x = defaultdict(float)
    for row in tqdm(data, desc="getting GI"):
        program = int(row[p])
        area = row[a]

        x[program] += area

    return x

def mb51():
    wb = xlwings.books["mb51.xlsx"]
    data = wb.sheets.active.range("A1").expand().value
    header = data.pop(0)

    p = header.index("Reference")
    d = header.index("Material Document")
    a = header.index("Quantity")

    x = defaultdict(list)
    for row in tqdm(data, desc="getting MB51"):
        program = int(row[p])
        doc = row[d]
        area = row[a]

        x[program].append((doc, -1 * area))

    return x

def nearest_target(ls, target):
    permutations = list()
    for i in range(1,len(ls)+1):
        permutations.extend(combinations(ls, i))

    nearest_sum = 1_000_000.0
    result = None
    for p in permutations:
        s = abs(target - sum([x[1] for x in p]))
        if s < nearest_sum:
            nearest_sum = s
            result = list([x[0] for x in p])

    return result, nearest_sum


a = gi()
b = mb51()

rev = list()
for program, target in tqdm(a.items(), desc="getting docs"):
    if program not in b:
        tqdm.write(f"Program {program} not in MB51")
    else:
        docs, s = nearest_target(b[program], target)
        rev.extend(docs)

print(tabulate([[x] for x in rev]))
with open("temp/docs.txt", 'w') as f:
    f.writelines("\n".join(rev))
