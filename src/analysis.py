from argparse import ArgumentParser
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from itertools import groupby
import logging
import os
import re
from typing import Tuple, List, Dict
import pyperclip
from tqdm import tqdm
from types import SimpleNamespace
import xlwings

from mb51 import Mb51, AnalysisMatch
from lib.db import SndbConnection

"""
# Process
- parse MB51 from SAP
    - mark any items that have an ID in reference as issued (1-1)
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
2) closest area
3) closest timestamp
"""

# add custom log level "TRACE"
TRACE = 5
logging.addLevelName(TRACE, "TRACE")


def trace(self, message, *args, **kw):
    self.log(TRACE, message, *args, **kw)


logging.Logger.trace = trace

# configure logging
logging.basicConfig(
    format="%(levelname)s: %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)


# https://stackoverflow.com/a/287944
class bcolors:
    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"


@dataclass
class AnalysisRowUpdate:
    """
    Row that has order/doc and consumption matched and needs updated
    """

    sapref: int
    consumption: float


@dataclass
class CompleteAnalysisRow:
    """
    Row that is filled out, but does not need to be written on update
    """

    sapref: int


@dataclass
class NotMatchedAnalysisRow:
    """
    Row that is not yet matched
    """

    part: str
    matl: str
    timestamp: datetime
    qty: int
    area: float


@dataclass
class ParsedAnalysisRow:
    id: int
    part: str
    matl: str
    timestamp: datetime
    qty: int
    area: float
    sapref: int | None
    sapval: float | None

    def to_not_matched(self) -> NotMatchedAnalysisRow:
        return NotMatchedAnalysisRow(
            part=self.part,
            matl=self.matl,
            timestamp=self.timestamp,
            qty=self.qty,
            area=self.area,
        )


@dataclass
class NearestNeighborDistance:
    area: float
    timestamp: timedelta

    def __lt__(self, other):
        return self.timestamp < other.timestamp and self.area <= other.area


@dataclass
class Neighborhood:
    # idea:
    #   Build a 2d memo table (x=analysis, y=mb51) to track the nearest
    #   neighbor based on distance to timestamp and area.

    part: str
    qty: int
    matl: str
    analysis: Dict[int, ParsedAnalysisRow]
    mb51: List[AnalysisMatch]
    matrix: Dict[int, List[NearestNeighborDistance]]

    def __init__(self, part, qty, matl):
        self.part = part
        self.qty = qty
        self.matl = matl
        self.analysis = dict()
        self.mb51 = list()
        self.matrix = dict()

    def add_analysis(self, id: int, row: ParsedAnalysisRow):
        self.analysis[id] = row
        self.matrix[id] = list()

    def add_mb51(self, row: AnalysisMatch):
        self.mb51.append(row)
        for k, v in self.analysis.items():
            timediff = row.timestamp - v.timestamp
            if row.timestamp < v.timestamp:
                timediff = timedelta.max

            self.matrix[k].append(
                NearestNeighborDistance(
                    area=abs(row.area - v.area),
                    timestamp=timediff,
                )
            )

    def get_min(self) -> Tuple[int, AnalysisMatch]:
        # get minimum distance in matrix
        min_key = None
        min_dist = None
        min_dist_index = None
        for k, v in self.matrix.items():
            for i, d in enumerate(v):
                if not min_dist or d < min_dist:
                    min_key = k
                    min_dist = d
                    min_dist_index = i

        if not min_key:
            return None

        # remove row and column of minimum distance
        # del self.analysis[min_key]
        del self.matrix[min_key]
        mb51 = self.mb51.pop(min_dist_index)
        for k, v in self.matrix.items():
            v.pop(min_dist_index)

        # return analysis id, mb51 item
        return min_key, mb51

    def dump_updates(self):
        while 1:
            x = self.get_min()
            if x:
                yield x
            else:
                return StopIteration


class WeeklyAnalysis:
    rows: Dict[int, ParsedAnalysisRow | AnalysisRowUpdate]
    mb51: Mb51
    _wb: xlwings.Book
    _sheet: xlwings.Sheet
    _monday: date
    _header: SimpleNamespace

    def __init__(self, monday=None):
        self.rows = dict()
        self._monday = monday
        self._header = None
        self._wb = None
        self._sheet = None

        self.mb51 = None

    def __del__(self):
        self.workbook.close()

    @property
    def monday(self):
        if not self._monday:
            today = date.today()
            self._monday = today.replace(day=today.day - today.weekday())

        return self._monday

    @property
    def workbook(self):
        if not self._wb:
            self._wb = xlwings.Book(
                r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx"
            )

        return self._wb

    @property
    def sheet(self):
        if not self._sheet:
            self._sheet = self.workbook.sheets[self.monday.strftime("%Y-%m-%d")]

        return self._sheet

    def pull(self):
        sqlfile = os.path.join(
            os.path.dirname(__file__), "sql", "get_analysis_data.sql"
        )

        # create sheet if it does not exist
        wb = self.workbook
        if self.monday() not in wb.sheet_names:
            wb.sheets["template"].copy(before=wb.sheets["Issues"], name=self.monday())

        # pull data if sheet is empty
        if self.sheet.range("A2").value is None:
            data = SndbConnection().query_from_sql_file(sqlfile)
            self.sheet.range("A2").value = [list(row) for row in data]
        else:
            print(
                "\033[91m Sheet {} already exists and has data\033[00m".format(
                    self.monday
                )
            )

        # copy parts and materials to clipboard
        self.get_not_matched()

    def get_not_matched(self):
        if not self.rows:
            self.parse_sheet()

        not_matched = list()
        dates = []
        for row in self.rows.values():
            match row:
                case ParsedAnalysisRow() if row.sapval is None:
                    not_matched += [row.part, row.matl]
                    dates.append(row.timestamp)
                case NotMatchedAnalysisRow():
                    not_matched += [row.part, row.matl]
                    dates.append(row.timestamp)

        # load not-matched into clipboard
        if not_matched:
            ls = sorted(set(not_matched))
            logging.debug(
                "Parts and Materials not matched:\n{}\n~~~~~~~~~~~~~~~~~~".format(
                    "\n".join(ls)
                )
            )

            pyperclip.copy("\r\n".join(ls))
            print(
                "Parts and Materials copied to clipboard. Date range is is {} to {}".format(
                    min(dates).strftime("%m-%d-%Y"), max(dates).strftime("%m-%d-%Y")
                )
            )

    def match(self):
        self.parse_sheet()
        self.analyze()
        self.write_updates()
        self.get_not_matched()

    @property
    def header(self):
        if not self._header:
            aliases = dict(
                id="Id",
                timestamp="UpdateDate",
                part="Part",
                program="Program",
                qty="Qty",
                area="Area",
                matl="MaterialMaster",
                sapref="OrderOrDocument",
                sapval="SAPValue",
            )

            self._header = SimpleNamespace(max=0)
            row = self.sheet.range("A1").expand("right").value
            for k, v in aliases.items():
                index = row.index(v)
                self._header.max = max(self._header.max, index)
                setattr(self._header, k, index)

        return self._header

    def parse_sheet(self):
        h = self.header
        rng = (
            self.sheet.range((2, 1), (2, h.max + 1))
            .expand("down")
            .options(ndim=2)
            .value
        )
        rng = tqdm(
            enumerate(rng, start=2),
            desc="Parsing sheet {}".format(self.sheet),
            total=len(rng),
        )
        for i, row in rng:
            log.trace(row)

            sapref = row[h.sapref]
            sapval = row[h.sapval]
            if sapref:
                sapref = int(sapref)
            if sapval:
                sapval = float(sapval)

            self.rows[i] = ParsedAnalysisRow(
                id=i,
                part=row[h.part],
                matl=row[h.matl],
                timestamp=row[h.timestamp],
                qty=int(row[h.qty]),
                area=row[h.area],
                sapref=sapref,
                sapval=sapval,
            )

    def analyze(self):
        self.mb51 = Mb51()

        # easy matches
        for k, row in self.rows.items():
            # all rows should be of type ParsedAnalysisRow
            assert isinstance(row, ParsedAnalysisRow), "Row is not ParsedAnalysisRow"

            match (row.sapref, row.sapval):
                # no order/doc -> needs matched
                case (None, None):
                    # try to match by ID, in case it's an issued item
                    by_id = self.mb51.get_by_id(row.id)
                    if by_id:
                        self.update(i, by_id.doc, by_id.area)

                    # set to be matched using nearest neighbor
                    else:
                        self.rows[k] = row.to_not_matched()

                # order/doc was manually tagged
                case (sapref, None):
                    area = self.mb51.get_area(int(sapref))
                    if area:
                        self.update(i, sapref, area)
                    else:
                        log.info("Order/Document `{}` not found".format(sapref))

                # already matched
                case (sapref, _):
                    self.rows[k] = CompleteAnalysisRow(sapref)
                    self.mb51.remove(sapref)

        log.debug("MB51 listing")
        log.debug("==============")
        for v in self.mb51.rows.values():
            log.debug(v)
        log.debug("==============")
        log.debug("Not-Matched listing")
        log.debug("==============")
        for v in self.rows.values():
            match v:
                case NotMatchedAnalysisRow(part, matl, timestamp, qty, area):
                    log.debug(v)
        log.debug("==============")

        # analyze
        key = lambda r: (r.part, r.qty, r.matl)
        neighborhoods = dict()
        for k, r in self.rows.items():
            match r:
                case NotMatchedAnalysisRow(part, matl, timestamp, qty, area):
                    key = (part, qty, matl)
                    if key not in neighborhoods:
                        neighborhoods[key] = Neighborhood(
                            part=part,
                            qty=qty,
                            matl=matl,
                        )

                    neighborhoods[key].add_analysis(k, r)
        for key, group in neighborhoods.items():
            for x in self.mb51.get_neighborhood(*key):
                group.add_mb51(x)

            log.debug(key)
            for k, x in group.analysis.items():
                log.debug("\t-> (%d), %s", k, x)
            for o in group.mb51:
                log.debug("\t-| %s", 0)
            for id, order in group.dump_updates():
                if log.level <= logging.DEBUG:
                    a = self.rows[id]
                    from_ts = a.timestamp.strftime("%Y-%m-%d %H:%M:%S")
                    to_ts = order.timestamp.strftime("%Y-%m-%d %H:%M:%S")
                    from_area = a.area
                    to_area = order.area
                    log.debug(
                        "\t<- {}({}) {} | {}{}".format(
                            bcolors.FAIL, id, from_ts, from_area, bcolors.ENDC
                        )
                    )
                    log.debug(
                        "\t   {}({}) {} | {}{}".format(
                            bcolors.OKGREEN, id, to_ts, to_area, bcolors.ENDC
                        )
                    )

                self.update(id, order.order, order.area)

    def write_updates(self):
        # calculate updates
        start = 0
        updates = dict()
        for k, item in self.rows.items():
            match item:
                case AnalysisRowUpdate(sapref, area):
                    if start == 0:
                        start = k
                        updates[start] = list()

                    updates[start].append([sapref, area])

        # write updates
        update_count = 0
        for start, rows in updates.items():
            self.sheet.range((start, self.header.sapref + 1)).value = rows
            update_count += len(rows)

        log.info("%d Rows updated", update_count)
        self.workbook.save()

    def update(self, row_id: int, order_or_doc: int, consumption: float):
        self.rows[row_id] = AnalysisRowUpdate(order_or_doc, consumption)
        self.mb51.remove(order_or_doc)

    def print(self):
        for k, row in self.rows.items():
            print(k, "->", row)

    def print_mb51(self):
        self.mb51.print()


if __name__ == "__main__":
    parser = ArgumentParser()
    parser.add_argument("-p", "--pull", action="store_true", help="get the data")
    parser.add_argument(
        "-a", "--analyze", action="store_true", help="fill in MB51 data"
    )
    parser.add_argument(
        "-n",
        "--not-matched",
        action="store_true",
        help="get materials that are not matched",
    )
    parser.add_argument(
        "--monday", type=str, default=None, help="Monday date operate on"
    )
    parser.add_argument(
        "-v", "--verbose", action="count", default=3, help="make the script more chatty"
    )
    parser.add_argument(
        "-q", "--quiet", action="count", default=0, help="make the script less chatty"
    )
    parser.add_argument(
        "-s",
        "--silence",
        action="store_true",
        help="I'm not interested in talking today",
    )
    args = parser.parse_args()

    if args.silence:
        args.verbose = 0

    args.verbose -= args.quiet
    match args.verbose:
        case i if i < 1:
            log.setLevel(logging.CRITICAL)
        case 1:
            log.setLevel(logging.ERROR)
        case 2:
            log.setLevel(logging.WARNING)
        case 3:
            log.setLevel(logging.INFO)
        case 4:
            log.setLevel(logging.DEBUG)
        case i if i > 4:
            log.setLevel(TRACE)

    if args.pull:
        WeeklyAnalysis(monday=args.monday).pull()
    elif args.analyze:
        WeeklyAnalysis(monday=args.monday).match()
    elif args.not_matched:
        WeeklyAnalysis(monday=args.monday).get_not_matched()
    else:
        print("No action specified")
