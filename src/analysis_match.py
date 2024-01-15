
from dataclasses import dataclass
from datetime import datetime, timedelta, date
import os
import re
from tqdm import tqdm
import xlwings

from lib.parsers import SheetParser

inbox_pattern = re.compile(r"Planned order not found for (\d{7}[a-zA-Z]-[\w-]+), (D-\d{7}-\d{5}), ([\d,]+).000, Sigmanest Program:([\d-]+)")

class Header:
    id = 0
    timestamp = 1
    part = 2
    program = 3
    qty = 4
    area = 5
    loc = 6
    mm = 7
    wbs = 8
    plant = 9
    order = 10
    sapval = 11
    notes = 12

@dataclass
class InboxError:
    part: str
    wbs: str
    qty: int
    program: str

@dataclass
class ProductionOrder:
    part: str
    order: int
    qty: int

@dataclass
class Mb51Item:
    matl: str
    order: ProductionOrder
    timestamp: datetime
    area: float

@dataclass
class IssueItem:
    matl: str
    doc: str
    timestamp: datetime
    prog: str
    area: float

@dataclass
class AnalysisRow:
    id: str
    part: str
    matl: str
    loc: str
    wbs: str
    plant: str
    timestamp: datetime
    prog: str
    qty: int
    area: float

    def parse(row):
        def numstr(val):
            if type(val) in (int, float):
                return str(int(val))
            return val

        return AnalysisRow(
            id=numstr(row[Header.id]),
            part=row[Header.part].upper(),
            matl=row[Header.mm],
            loc=row[Header.loc],
            wbs=row[Header.wbs],
            plant=row[Header.plant],
            timestamp=row[Header.timestamp],
            prog=numstr(row[Header.program]),
            qty=row[Header.qty],
            area=row[Header.area],
        )

def main():
    inbox = []
    if os.path.exists("./inbox.txt"):
        with open("inbox.txt") as inbx:
            for line in inbx.readlines():
                match = inbox_pattern.match(line.strip())
                if match:
                    vals = list(match.groups())
                    vals[2] = int(vals[2])
                    inbox.append(InboxError(*vals))

    cnf, issue = parse_mb51()

    strategies = range(0, 4)
    def get_consumption(analysisRow, strategy=0):
        order_or_doc = None
        area = None

        match strategy:
            case 0:
                # direct matches
                for i, item in enumerate(cnf):
                    area_match = abs(item.area - analysisRow.area) < .001
                    if item.order.part == analysisRow.part and item.matl == analysisRow.matl and area_match and item.order.qty == analysisRow.qty and item.timestamp > analysisRow.timestamp:
                        order_or_doc = item.order.order
                        area = item.area
                        cnf.pop(i)
                        
                        return order_or_doc, area
                    
            case 1:
                # match with a wider range for area
                for i, item in enumerate(cnf):
                    area_match = abs(item.area - analysisRow.area) < 100
                    if item.order.part == analysisRow.part and item.matl == analysisRow.matl and area_match and item.order.qty == analysisRow.qty and item.timestamp > analysisRow.timestamp:
                        order_or_doc = item.order.order
                        area = item.area
                        cnf.pop(i)
                        
                        return order_or_doc, area
                    
            case 2:
                # match for issue items
                for i, item in enumerate(issue):
                    area_match = abs(item.area - analysisRow.area) < .001
                    if item.matl == analysisRow.matl and (item.prog == analysisRow.prog or item.prog == analysisRow.id) and area_match and item.timestamp > analysisRow.timestamp:
                        order_or_doc = item.doc
                        area = item.area
                        issue.pop(i)
                        
                        return order_or_doc, area
                    

            case 3:
                # direct matches, dates not right (can happen with COGI clearing)
                for i, item in enumerate(cnf):
                    area_match = abs(item.area - analysisRow.area) < .001
                    if item.order.part == analysisRow.part and item.matl == analysisRow.matl and area_match and item.order.qty == analysisRow.qty and (analysisRow.timestamp - item.timestamp).days == 0:
                        order_or_doc = item.order.order
                        area = item.area
                        cnf.pop(i)
                        
                        return order_or_doc, area

            case 'inbox':
                # direct matches, dates not right (can happen with COGI clearing)
                for i, item in enumerate(inbox):
                    if item.part == analysisRow.part and item.qty == analysisRow.qty and item.program == analysisRow.prog:
                        inbox.pop(i)
                        return True
                else:
                    return False

            case _:
                pass
                
        return None

    today = date.today()
    monday = today.replace(day=today.day-today.weekday())

    wb = xlwings.Book(r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx")
    sheet = wb.sheets[monday.strftime("%Y-%m-%d")]
    # sheet = wb.sheets["2023-12-25"]
    data = list()
    vals = sheet.range("A2:J2").expand('down').value
    print("filling sheet", sheet.name)
    for r, row in tqdm(enumerate(vals, start=2), desc="Parsing Analysis", total=len(vals)):
        if sheet.range((r, Header.order+1)).value or sheet.range((r, Header.sapval+1)).value:
            data.append(None)
            continue

        data.append( AnalysisRow.parse(row) )


    updates_made = 0
    for strategy in strategies:
        for r, row in tqdm(enumerate(data, start=2), desc=f"Setting Data<strategy:{strategy}>", total=len(data)):
            if row is None:
                continue

            match = get_consumption(row, strategy)
            if match:
                sheet.range((r, Header.order+1), (r, Header.sapval+1)).value = match
                sheet.range((r, Header.order+1), (r, Header.sapval+1)).color = "#F4128B"
                data[r-2] = None
                updates_made += 1

    for r, row in tqdm(enumerate(data, start=2), desc=f"Setting Data<strategy:inbox>", total=len(data)):
        if row is None:
            continue

        is_in_inbox = get_consumption(row, 'inbox')
        if is_in_inbox:
            sheet.range((r, Header.sapval+2)).value = 'Inbox error'
            data[r-2] = None
            updates_made += 1

    print(f"{updates_made} rows were updated")
    wb.close()
        


def parse_mb51() -> dict[str, Mb51Item]:

    cnf, issue = list(), list()

    if len(xlwings.apps) == 0:
        app = xlwings.App()
        wb = app.books.open(r"C:\Users\PMiller1\Documents\SAP\SAP GUI\mb51.xlsx")
        app.books['Book1'].close()
    wb = SheetParser(wb='mb51.xlsx')

    orders = parse_cohv(skip_if_not_open=True)

    # TODO: change parser so that we don't use last qty column.
    #   This will remove the need to convert from FT2,
    #   which introduces a mismatch due to conversion

    sort_fn = lambda r: (r.type, r.matl)
    skip_fn = lambda x: not x.matl

    for row in sorted(wb.parse_sheet(with_progress=True, skip_if=skip_fn), key=sort_fn):
        # row.qty is negative for a consumption
        time = timedelta(days=row.time)
        timestamp = row.date+time
        area = -1 * row.qty

        if row.uom == "FT2":
            area *= 144
            
        match row.type:
            case '101' if row.loc == 'PROD':
                orders[row.order] = ProductionOrder(row.matl, row.order, row.qty)
            case '201' | '221' if row.program is not None:
                # issue to cost center
                if type(row.program) in (int, float):
                    row.program = str(int(row.program))
                issue.append( IssueItem(row.matl, row.document, timestamp, row.program, area) )
            case '261':
                # issue to order
                try:
                    cnf.append( Mb51Item(row.matl, orders[row.order], timestamp, area) )
                except KeyError:
                    if 'BATCH' in row.user:
                        tqdm.write(f"part not found for order: {row.order}")
            case _:
                continue

    wb.workbook.close()

    return cnf, issue


def parse_cohv(skip_if_not_open = False) -> dict[str, (str, int)]:
    orders = dict()

    if 'cohv.xlsx' in xlwings.books and skip_if_not_open == False:
        for row in SheetParser(wb='cohv.xlsx').parse_sheet(with_progress=True):
            orders[row.order] = ProductionOrder(row.part, row.order, row.qty)

    return orders


def diagnose():
    today = date.today()
    monday = today.replace(day=today.day-today.weekday())

    print("MB51 data")
    cnf, issue = parse_mb51(parse_cohv(skip_if_not_open=True))
    for x in cnf:
        if x.matl == '1220203A01-04003':
            print(x)

    wb = xlwings.Book(r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx")
    sheet = wb.sheets[monday.strftime("%Y-%m-%d")]
    sheet = wb.sheets["2023-12-25"]
    print("Analysis data")
    for row in sheet.range("A2:J2").expand('down').value:
        row = AnalysisRow.parse(row)
        if row.matl == '1220203A01-04003':
            print(row)

if __name__ == "__main__":
    main()
    # diagnose()
