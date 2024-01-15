
import os

from collections import defaultdict
from tabulate import tabulate
from tqdm import tqdm

from lib.db import SndbConnection
from lib.parsers import SheetParser

os.system('cls')
data = SheetParser().parse_sheet(with_progress=True)

sap = defaultdict(float)
for row in data:
    if row.matl and ( '-03' in row.matl or '-04' in row.matl ):
        if row.units == 'FT2':
            row.qty *= 144
        elif row.units == 'M2':
            row.qty *= 1550
        sap[row.matl] += row.qty


table = ["MM,SAP,Sigmanest,Diff".split(',')]
possible_not_cnf = list()
with SndbConnection() as db:
    c = db.cursor
    for matl, sap_qty in tqdm(sap.items(), desc="Querying database"):
        c.execute("""
                  SELECT ISNULL( SUM(Area), 0 )
                  FROM Stock WHERE PrimeCode=? and SheetName LIKE 'W%'
                  """, matl)
        sn_qty = c.fetchone()[0]

        if sap_qty-sn_qty > 1000.0:
            table.append([matl, sap_qty, sn_qty, sap_qty-sn_qty])
            c.execute("""
                      SELECT PIPArchive.PartName
                      FROM PIPArchive
                      INNER JOIN StockHistory
                        ON  StockHistory.ProgramName=PIPArchive.ProgramName
                        AND StockHistory.SheetName  =PIPArchive.SheetName
                      WHERE StockHistory.PrimeCode=? AND PIPArchive.TransType='SN102'
                      """, matl)
            possible_not_cnf.extend([x[0].replace('_', '-', 1) for x in c.fetchall()])

print(tabulate(table, headers="firstrow"))

with open('parts.txt', 'w') as f:
    f.write('\n'.join(possible_not_cnf))
