from argparse import ArgumentParser
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from itertools import groupby
import logging
import os
import re
from typing import Iterator, List, Dict
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
class ParsedAnalysisRow:
    part: str
    matl: str
    timestamp: datetime
    qty: int
    area: float


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

    def get_min(self) -> (int, AnalysisMatch):
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

    def __init__(self, monday=None):
        self.rows = dict()
        self._monday = monday

        self.mb51 = Mb51()

    def __del__(self):
        self.wb.close()

    @property
    def monday():
        if not self._monday:
            today = date.today()
            self._monday = today.replace(day=today.day - today.weekday())

        return self._monday

    @property
    def workbook(self):
        return xlwings.Book(
            r"C:\Users\PMiller1\OneDrive - high.net\inventory\InventoryAnalysis\2023_WeeklyAnalysis.xlsx"
        )

    @property
    def sheet(self):
        return self.workbook.sheets[self.monday.strftime("%Y-%m-%d")]

    def pull(self):
        sqlfile = os.path.join(
            os.path.dirname(__file__), "sql", "get_analysis_data.sql"
        )

        # create sheet if it does not exist
        if self.monday() not in wb.sheet_names:
            wb.sheets["template"].copy(before=wb.sheets["Issues"], name=self.monday())

        # pull data if sheet is empty
        if self.sheet.range("A2").value is None:
            data = SndbConnection().query_from_sql_file(sqlfile)
            self.sheet.range("A2").value = [list(row) for row in data]
        else:
            print(
                "\033[91m Sheet {} already exists and has data\033[00m".format(
                    sheet_name
                )
            )

        # copy parts and materials to clipboard
        self.get_not_matched()

    def get_not_matched(self):
        if not self.rows:
            self.parse_sheet()

        not_matched = list()
        earliest_date = datetime.max
        for item in self.rows.values():
            match item:
                case ParsedAnalysisRow(part, matl, timestamp, qty, area):
                    not_matched += [part, matl]
                    earliest_date = min(earliest_date, timestamp)

        # load not-matched into clipboard
        if not_matched:
            pyperclip.copy("\r\n".join(sorted(set(not_matched))))
            print(
                "Parts and Materials copied to clipboard. Earliest date is {}".format(
                    earliest_date.strftime("%m-%d-%Y")
                )
            )

    def match(self):
        self.parse_sheet()
        self.analyze()
        self.write_updates()
        self.get_not_matched()

    def parse_sheet():
        vals = self.sheet.range("A2:J2").expand("down").value
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

        header = SimpleNamespace()
        row = self.sheet.range("A1").expand("right").value
        for k, v in aliases.items():
            setattr(header, k, row.index(v))

        rng = (
            sheet.range((2, 1), (2, len(row) + 1)).expand("down").options(ndim=2).value
        )
        rng = tqdm(
            enumerate(rng, start=2),
            desc="Parsing sheet {}".format(sheet),
            total=len(rng),
        )
        for i, row in rng:
            match (row[header.sapref], row[header.sapval]):
                case (None, None):
                    by_id = self.mb51.get_by_id(int(row[header.id]))
                    if by_id:
                        self.update(i, by_id.doc, by_id.area)
                        continue

                    self.rows[i] = ParsedAnalysisRow(
                        part=row[header.part],
                        matl=row[header.matl],
                        timestamp=row[header.timestamp],
                        qty=int(row[header.qty]),
                        area=row[header.area],
                    )

                case (sapref, None):
                    area = self.mb51.get_area(int(sapref))
                    if area:
                        self.update(key, sapref, area)
                    else:
                        tqdm.write("Order/Document `{}` not found".format(sapref))

                case (sapref, _):
                    self.mb51.remove(sapref)

    def analyze(self):
        # analyze
        key = lambda r: (r.part, r.qty, r.matl)
        neighborhoods = dict()
        for k, r in self.rows.items():
            match r:
                case ParsedAnalysisRow(part, matl, timestamp, qty, area):
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

                self.rows[id] = AnalysisRowUpdate(order.order, order.area)
                self.mb51.remove(order.order)

    def write_updates(self):
        # calculate updates
        start = 0
        updates = dict()
        not_matched = list()
        earliest_date = datetime.max
        for k, item in self.rows.items():
            match item:
                case AnalysisRowUpdate(sapref, area):
                    if start == 0:
                        start = k
                        updates[start] = list()

                    updates[start].append([sapref, area])

                case ParsedAnalysisRow(part, matl, timestamp, qty, area):
                    # reset updates start counter
                    start = 0

                    not_matched += [part, matl]
                    earliest_date = min(earliest_date, timestamp)

        # write updates
        update_count = 0
        for start, updates in updates.items():
            sheet.range((start, header.sapref + 1)).value = updates
            update_count += len(updates)

        log.info("%d Rows updated", update_count)
        wb.save()

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
        "-v", "--verbose", action="count", help="make the script more chatty"
    )
    parser.add_argument(
        "-q", "--quiet", action="count", help="make the script less chatty"
    )
    parser.add_argument(
        "-s",
        "--silence",
        action="store_true",
        help="I'm not interested in talking today",
    )
    args = parser.parse_args()

    if args.silence:
        verbose = -1
    else:
        verbose = 3 + args.verbose - args.quiet
    match verbose:
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
            log.setLevel(logging.TRACE)

    if args.pull:
        WeeklyAnalysis(monday=args.monday).pull()
    elif args.analyze:
        WeeklyAnalysis(monday=args.monday).match()
    else:
        print("No action specified")
