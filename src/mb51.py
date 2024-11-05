
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Iterator, Dict, List
from tqdm import tqdm
from types import SimpleNamespace

import xlwings

from analysis_match import AnalysisRowUpdate


def numstr(val) -> str:
    if type(val) in (int, float):
        return str(int(val))
    return val

@dataclass
class IssueItem:
    matl: str
    doc: str
    timestamp: datetime
    ref: str
    area: float

    def to_update(self) -> AnalysisRowUpdate:
        return AnalysisRowUpdate(sapref=self.doc, consumption=self.area)

@dataclass
class CnfItem:
    matl: str
    timestamp: datetime
    area: float

@dataclass
class ProductionOrder:
    part: str
    order: int
    qty: int
    material: CnfItem | None

    def to_update(self) -> AnalysisRowUpdate:
        assert self.material is None

        return AnalysisRowUpdate(sapref=self.order, consumption=self.material.area)

Mb51Item = CnfItem | IssueItem

@dataclass
class Mb51ParsedRow:
    doc: int;
    mvmt: str;
    matl: str;
    qty: float;
    order: str;
    timestamp: datetime;
    ref: str;

    def to_order(self, cnf=None) -> ProductionOrder:
        return ProductionOrder(
            part=self.matl, order=self.order, qty=self.qty, material=cnf
        )

    def to_issued(self) -> IssueItem:
        return IssueItem(
            matl=self.matl,
            doc=self.doc,
            timestamp=self.timestamp,
            ref=self.ref,
            # row.qty is negative for a consumption
            area=-1*self.qty
        )
    
    def to_cnf(self) -> CnfItem:
        return CnfItem(
            matl=self.matl,
            timestamp=self.timestamp,
            # row.qty is negative for a consumption
            area=-1*self.qty
        )

class Mb51:
    issued: Dict[str, IssueItem]
    cnf: Dict[str, ProductionOrder]

    def __init__(self) -> None:
        self.issued = dict()
        self.cnf = dict()

        self.parse_sheet()

    def print(self) -> None:
        print("Issued:")
        for x in self.issued.values():
            print('\t', x)

        print("Cnf:")
        for x in self.cnf.values():
            print('\t', x)

    def commit_order(self, order: str) -> AnalysisRowUpdate:
        ref = numstr(ref)

        result = self.cnf[order].to_update()
        del self.cnf[order]

        return result

    def commit_issued(self, ref: str | int | float) -> AnalysisRowUpdate:
        ref = numstr(ref)

        result = self.issued[ref].to_update()
        del self.issued[ref]

        return result

    def get_orders_by_part(self, part: str) -> Iterator[ProductionOrder]:
        for v in self.cnf.values():
            if v.material == part:
                yield v

    def get_issued_by_mm(self, mm: str) -> Iterator[ProductionOrder]:
        for v in self.issued.values():
            if v.matl == mm:
                yield v

    def parse_sheet(self) -> Iterator[Mb51ParsedRow]:
        if len(xlwings.apps) == 0:
            app = xlwings.App()
            wb = app.books.open(r"C:\Users\PMiller1\Documents\SAP\SAP GUI\mb51.xlsx")
            app.books['Book1'].close()
        else:
            wb = xlwings.books['mb51.xlsx']
        sheet = wb.sheets.active

        aliases = dict(
            matl = "Material",
            uom = "Unit of Entry",
            qty = "Qty in unit of entry",
            type = "Movement type",
            loc = "Storage Location",
            plant = "Plant",
            order = "Order",
            date = "Posting Date",
            time = "Time of Entry",
            id = "Reference",
            document = "Material Document",
            user = "User Name",
        )

        header = SimpleNamespace()
        row = sheet.range("A1").expand('right').value
        for k, v in aliases.items():
            setattr(header, k, row.index(v))

        rng = sheet.range((2, 1), (2, len(row) + 1)).expand('down').options(ndim=2).value
        rng = tqdm(rng, desc='Parsing sheet {}'.format(sheet), total=len(rng))

        cnf: Dict[str, CnfItem] = dict()
        orders: List[Mb51ParsedRow] = list()
        for row in rng:
            parsed = Mb51ParsedRow(
                doc=row[header.document],
                mvmt=row[header.type],
                matl=row[header.matl],
                qty=row[header.qty],
                order=row[header.order],
                timestamp=row[header.date] + timedelta(days=row[header.time]),
                ref=row[header.id],
            )

            # TODO: change parser so that we don't use last qty column.
            #   This will remove the need to convert from FT2,
            #   which introduces a mismatch due to conversion
            if row[header.uom] == "FT2":
                parsed.qty *= 144
                
            match parsed.mvmt:
                case '101' if row[header.loc] == 'PROD':
                    orders.append(parsed)
                case '201' | '221' if parsed.ref is not None:
                    # issue to cost center or job
                    self.issued[parsed.ref] = parsed.to_issued()
                case '261' if row[header.uom] != 'EA':
                    # issue to order
                    cnf[parsed.order] = parsed.to_cnf()
                case _:
                    continue

        # wb.close()

        for parsed in orders:
            self.cnf[parsed.order] = parsed.to_order(cnf.get(parsed.order, None))
