from dataclasses import dataclass
from datetime import datetime, timedelta, date
import re
from typing import List
from tqdm import tqdm
from types import SimpleNamespace
import xlwings

from mb51 import numstr, Mb51

inbox_pattern = re.compile(
    r"Planned order not found for (\d{7}[a-zA-Z]-[\w-]+), (D-\d{7}-\d{5}), ([\d,]+).000, Sigmanest Program:([\d-]+)"
)


class Header:
    id = (0,)
    timestamp = (1,)
    part = (2,)
    program = (3,)
    qty = (4,)
    area = (5,)
    loc = (6,)
    mm = (7,)
    wbs = (8,)
    plant = (9,)
    order = (10,)
    sapval = (11,)
    notes = (12,)


@dataclass
class AnalysisRow:
    id: int
    part: str


@dataclass
class AnalysisRowUpdate(AnalysisRow):
    sapref: str
    consumption: float


@dataclass
class ParsedAnalysisRow(AnalysisRow):
    # TODO: ParsedAnalysisRow -> NeedsMatch | IsMatched
    part: str
    matl: str
    loc: str
    wbs: str
    plant: str
    timestamp: datetime
    program: str
    qty: int
    area: float
    sapref: str
    consumption: float

    def parse(row):
        return AnalysisRow(
            id=int(row[Header.id]),
            part=row[Header.part].upper(),
            matl=row[Header.mm],
            loc=row[Header.loc],
            wbs=row[Header.wbs],
            plant=row[Header.plant],
            timestamp=row[Header.timestamp],
            program=numstr(row[Header.program]),
            qty=row[Header.qty],
            area=row[Header.area],
            sapref=row[Header.order],
            consumption=row[Header.sapval],
        )

    def to_update(self) -> AnalysisRowUpdate:
        return AnalysisRowUpdate(sapref=self.sapref, consumption=self.consumption)


@dataclass
class ProductionOrder:
    # TODO: ProductionOrder -> NeedsMatch | IsMatched
    part: str
    order: int
    qty: int
    consumption: ConsumptionItem | None


@dataclass
class ConsumptionItem:
    matl: str
    timestamp: datetime
    area: float


@dataclass
class IssueItem:
    # TODO: IssueItem -> NeedsMatch | IsMatched
    matl: str
    doc: str
    timestamp: datetime
    prog: str
    area: float


def main() -> None:
    """
    # Process
    - parse MB51 from SAP
        - mark any items that have an ID in reference as issued (1-1)
        - mark any items that have a program in referece as issued (1-+)
        - match up raw consumption with production orders
    - parse weekly analysis
        - if item has Order/Doc:
            - if item has Area -> mark as committed in MB51 list
            - else (no Area) -> update area and mark as committed in MB51 list
    - match up weekly analysis with MB51
        - directly match issued items
        - create neighborhoods of connected data sets
            - group by part and material
            - nearest neighbor
    - write changes to analysis
    - load not-matched into clipboard

    # Nearest Neighbor Weighting
    1) (issue item) ID match
    2) (issue item) Program match
        - closest area
    3) closest area
    4) closest timestamp
    """

    mb51 = Mb51()

    today = date.today()
    monday = today.replace(day=today.day - today.weekday())

    wb = xlwings.Book(
        r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx"
    )
    sheet = wb.sheets[monday.strftime("%Y-%m-%d")]
    # sheet = wb.sheets["2023-12-25"]

    rows: List[AnalysisRow] = list()
    vals = sheet.range("A2:J2").expand("down").value

    print("filling sheet", sheet.name)
    for r, row in tqdm(
        enumerate(vals, start=2), desc="Parsing Analysis", total=len(vals)
    ):
        parsed = AnalysisRow.parse(row)

        if parsed.sapref and not parsed.consumption:
            if parsed.sapref in mb51.cnf:
                rows.append(AnalysisRow(id=r, data=mb51.commit_order(parsed.sapref)))
            else:
                rows.append(AnalysisRow(id=r, data=None))

        elif parsed.consumption:
            rows.append(AnalysisRow(id=r, data=None))

        rows.append(AnalysisRow(id=r, data=parsed))

    with open("temp/parsed.txt", "w") as f:
        f.write("\n".join([str(r) for r in rows if r]))
    with open("temp/sap.txt", "w") as f:
        f.write("\n".join([str(r) for r in mb51.issued]))
        f.write("\n".join([str(r) for r in mb51.cnf]))

    # TODO: make matches
    for row in rows:
        match row.data:
            case ParsedAnalysisRow():
                if row.data.id in mb51.issued:
                    row = mb51.commit_issued(row.data.id)
                # elif row.data.program in mb51.issued:
                #     row = mb51.commit_issued(row.data.program)

                # TODO: nearest neighbor
            case _:
                pass
        # issued item

    with open("temp/updates.txt", "w") as f:
        f.write("\n".join([str(r) for r in rows if r]))

    updates_made = 0
    for row in rows:
        match row.data:
            case AnalysisRowUpdate(sapref, consumption):
                # sheet.range((row.id, Header.order+1)).value = [sapref, consumption]
                updates_made += 1
            case _:
                pass

    print(f"{updates_made} rows were updated")
    wb.save()
    wb.close()

    # TODO: load not-matched into clipboard


if __name__ == "__main__":
    main()
