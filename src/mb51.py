from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, Iterator
from tqdm import tqdm
from types import SimpleNamespace

import xlwings


@dataclass
class ConsumptionItem:
    matl: str
    timestamp: datetime
    area: float


@dataclass
class IssueItem:
    id: int
    matl: str
    doc: int
    timestamp: datetime
    area: float


@dataclass
class ProductionOrder:
    part: str
    order: int
    qty: int
    consumption: ConsumptionItem | None

    def to_match(self):
        assert (
            self.consumption is not None
        ), "Cannot coerce a non-consumption to AnalysisMatch"

        return AnalysisMatch(
            order=self.order,
            timestamp=self.consumption.timestamp,
            area=self.consumption.area,
        )


@dataclass
class AnalysisMatch:
    order: int
    timestamp: datetime
    area: float


@dataclass
class Mb51ParsedRow:
    doc: int
    mvmt: str
    matl: str
    qty: float
    order: str | None
    timestamp: datetime
    ref: int | None

    def __init__(self, doc, mvmt, matl, qty, order, timestamp, ref):
        self.doc = int(doc)
        self.mvmt = mvmt
        self.matl = matl
        self.qty = qty
        self.timestamp = timestamp

        if order:
            self.order = int(order)

        if ref:
            try:
                self.ref = int(ref)
            except ValueError:
                pass

    def to_consumption(self) -> ConsumptionItem:
        return ConsumptionItem(
            matl=self.matl,
            timestamp=self.timestamp,
            area=self.qty * -1,
        )

    def to_issued(self) -> IssueItem:
        return IssueItem(
            id=self.ref,
            matl=self.matl,
            doc=self.doc,
            timestamp=self.timestamp,
            area=self.qty * -1,
        )

    def to_order(self) -> ProductionOrder:
        assert self.order is not None, "Cannot coerce a non-order to ProductionOrder"

        return ProductionOrder(
            part=self.matl,
            order=self.order,
            qty=int(self.qty),
            consumption=None,
        )


class Mb51:
    rows: Dict[int, ProductionOrder | IssueItem]
    _wb: xlwings.Book
    _sheet: xlwings.Sheet

    def __init__(self) -> None:
        self.rows = dict()
        self._wb = None
        self._sheet = None

        self.parse_sheet()

    def __del__(self):
        self.workbook.close()

    @property
    def workbook(self):
        if not self._wb:
            self._wb = xlwings.Book(
                r"C:\Users\PMiller1\Documents\SAP\SAP GUI\mb51.xlsx"
            )

        return self._wb

    @property
    def sheet(self):
        if not self._sheet:
            self._sheet = self.workbook.sheets[0]

        return self._sheet

    def parse_sheet(self):
        aliases = dict(
            matl="Material",
            uom="Unit of Entry",
            qty="Qty in unit of entry",
            type="Movement type",
            loc="Storage Location",
            plant="Plant",
            order="Order",
            date="Posting Date",
            time="Time of Entry",
            id="Reference",
            document="Material Document",
            user="User Name",
        )

        header = SimpleNamespace()
        row = self.sheet.range("A1").expand("right").value
        for k, v in aliases.items():
            setattr(header, k, row.index(v))

        rng = (
            self.sheet.range((2, 1), (2, len(row) + 1))
            .expand("down")
            .options(ndim=2)
            .value
        )
        rng = tqdm(rng, desc="Parsing sheet {}".format(self.sheet), total=len(rng))

        parse = lambda row: Mb51ParsedRow(
            doc=row[header.document],
            mvmt=row[header.type],
            matl=row[header.matl],
            qty=row[header.qty],
            order=row[header.order],
            timestamp=row[header.date] + timedelta(days=row[header.time]),
            ref=row[header.id],
        )

        consumption: Dict[int, ConsumptionItem] = dict()
        for row in rng:
            # TODO: change parser so that we don't use last qty column.
            #   This will remove the need to convert from FT2,
            #   which introduces a mismatch due to conversion
            if row[header.uom] == "FT2":
                row[header.qty] *= 144

            match row[header.type]:
                case "101" if row[header.loc] == "PROD" and row[
                    header.order
                ] is not None:
                    parsed = parse(row)
                    self.rows[parsed.order] = parsed.to_order()
                case "201" | "221" if row[header.id] is not None:
                    # issue to cost center or job
                    parsed = parse(row)
                    self.rows[parsed.doc] = parsed.to_issued()
                case "261" if row[header.uom] != "EA":
                    # issue to order
                    parsed = parse(row)
                    consumption[parsed.order] = parsed.to_consumption()
                case _:
                    pass

        for parsed in self.rows.values():
            match parsed:
                case ProductionOrder() if parsed.order in consumption:
                    parsed.consumption = consumption[parsed.order]
                    del consumption[parsed.order]
                case _:
                    pass

    def remove(self, order_or_doc):
        try:
            del self.rows[order_or_doc]
        except KeyError:
            pass

    def get_area(self, order_or_doc) -> float | None:
        match self.rows[order_or_doc]:
            case ProductionOrder(_, _, _, consumption):
                return consumption.area
            case IssueItem(_, _, _, _, area):
                return area
            case _:
                return None

    def get_by_id(self, id: int) -> IssueItem | None:
        for row in self.rows.values():
            if isinstance(row, IssueItem) and row.id == id:
                return row

        return None

    def print(self):
        for k, v in self.rows.items():
            print(k, "->", v)

    def get_neighborhood(
        self, part: str, qty: int, matl: str
    ) -> Iterator[AnalysisMatch]:
        for row in self.rows.values():
            match row:
                case (
                    ProductionOrder()
                ) if row.part == part and row.qty == qty and row.consumption and row.consumption.matl == matl:
                    yield row.to_match()
