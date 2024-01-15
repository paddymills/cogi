
import xlwings

from multiprocessing import Pool
from os import path, listdir, environ
from re import compile as regex
from types import SimpleNamespace
from tqdm import tqdm

aliases = SimpleNamespace()
aliases.matl = ("Material", "Material Number")
aliases.material = aliases.matl
aliases.part = aliases.matl
aliases.mark = aliases.matl
aliases.uom = ("Base Unit of Measure", "Unit of Entry")
aliases.qty = ("Qty in unit of entry", "Unrestricted", "Order quantity (GMEIN)", "Quantity")
aliases.shipment = ("Occurrence",)
aliases.type = ("Order Type", "Movement type")
aliases.loc = ("Storage Location",)
aliases.wbs = ("WBS Element", "Special stock number")
aliases.plant = ("Plant",)
aliases.order = ("Order",)
aliases.desc = ("Material description",)
aliases.date = ("Posting Date",)
aliases.time = ("Time of Entry",)
aliases.program = ("Reference",)
aliases.autoid = ("Reference",)
aliases.id = (*aliases.autoid, "Id")
aliases.cost_center = ("Cost Center",)
aliases.document = ("Material Document",)
aliases.user = ("User Name",)

BASE_SAP_DATA_FILES = r"\\hssieng\SNData\SimTrans\SAP Data Files"
SAP_OUTBOX = r"\\hiifileserv1\sigmanestprd\Archive"
SAP_EXPORTS = path.join(environ.get("USERPROFILE"), r"Documents\SAP\SAP GUI")

SCAN_PART = regex(r"(\d{3})([a-zA-Z])-\w+-\w+-(\w+)")
DATA_FILE_JOB = regex(r"S-(\d{7})")


def data_file_folder(folder_name):
    return path.join(BASE_SAP_DATA_FILES, folder_name)


def part_name(_part, _job):
    scan_match = SCAN_PART.match(_part)
    job_match = DATA_FILE_JOB.match(_job)
    if not (scan_match and job_match):
        return _part

    job_end, structure, part = scan_match.groups()
    job_without_structure = job_match.group(1)

    # if _part == "252B-1B-X-M318A":
    #     print("\nmatches:", scan_match.groups(), "|", job_match.groups())
    #     print("scan match:", bool(scan_match and job_match))
    #     print("endswith match:", job_without_structure.endswith(job_end))
    #     print("partname:", "{}{}-{}".format(job_without_structure, structure, part))

    if not job_without_structure.endswith(job_end):
        return _part

    return "{}{}-{}".format(job_without_structure, structure, part)


class SheetParser:

    def __init__(self, wb=None, sheet=None):

        self._header = None
        self._sheet = None
        self._wb = wb

        if sheet:
            self.set_sheet(sheet)

    def set_sheet(self, sheet):

        if sheet is None:
            return

        elif sheet == 'active':
            self._sheet = xlwings.books.active.sheets.active
        
        elif type(sheet) is xlwings.Sheet:
            self._sheet = sheet

        elif type(sheet) in (str, int):
            self._sheet = xlwings.books.active.sheets[sheet]

        else:
            raise TypeError("sheet must be an xlwings sheet, string or integer")

    @property
    def header(self):

        if self._header is None:
            self.parse_header()

        return self._header
    
    @property
    def workbook(self):
        if type(self._wb) is str:
            if '.' not in self._wb:
                self._wb += ".xlsx"

            if xlwings.apps.count > 0 and self._wb in xlwings.books:
                self._wb = xlwings.books[self._wb]
            else:
                self._wb = xlwings.Book(path.join(SAP_EXPORTS, self._wb))

        else:
            self._wb = xlwings.books.active
        
        return self._wb

    @property
    def sheet(self):
        """
        returns sheet if set, otherwise active sheet

        this allows operation on active sheet without setting the stored sheet
        """

        if self._sheet is None:    
            self._sheet = self.workbook.sheets.active

        return self._sheet

    @property
    def max_col(self):

        return max(self.header.__dict__.values())

    @property
    def last_row(self):

        return self.sheet.range((2, self.header.matl + 1)).end('down').options(ndim=2).row

    @property
    def data_rng(self, start_row=2):
        return self.sheet.range((start_row, 1), (self.last_row, self.max_col + 1)).options(ndim=2)

    def parse_header(self, row=None):

        self._header = SimpleNamespace()

        if row is None:
            row = self.sheet.range("A1").expand('right').value

        for i, item in enumerate(row):
            # same as
            # `
            #   if item in aliases.matl:
            #       header.matl = i
            # `
            # just less code to maintain
            for k, v in aliases.__dict__.items():
                if item in v:
                    setattr(self._header, k, i)

        self._header_parsed = True

    def parse_row(self, row=None):
        res = SimpleNamespace()

        if row is None:
            row = 2
        if type(row) is int:
            row = self.sheet.range((row, 1)).expand('right').value

        for k, v in self.header.__dict__.items():
            # same as res.matl = row[header.matl]
            setattr(res, k, row[v])

        return res

    def parse_sheet(self, sheet=None, with_progress=False, skip_if=lambda: False):
        """
            convenience method to parse entire sheet
        """

        if sheet:
            self.set_sheet(sheet)

        if self.header is None:
            self.parse_header()

        rng = self.data_rng.value

        if with_progress:
            rng = tqdm(rng, desc='Parsing sheet {}'.format(self.sheet), total=len(rng))

        for row in rng:
            parsed = self.parse_row(row)

            if skip_if(parsed):
                continue

            yield parsed

class CnfFileParser:

    def __init__(self, processed_only=False, dir=None):
        self.ipart = SimpleNamespace(matl=0, qty=4, wbs=2, plant=11, job=1, program=12)
        self.imatl = SimpleNamespace(matl=6, qty=8, loc=10, wbs=7, plant=11)

        if dir is None:
            self.dirs = [
                data_file_folder("processed"),
                # SAP_OUTBOX,
            ]

            if not processed_only:
                for d in ('deleted files', '_temp', 'original'):
                    _dir = data_file_folder(d)
                    if path.exists(_dir):
                        self.dirs.append(_dir)
        elif type(dir) is list:
            self.dirs = dir
        else:
            assert type(dir) is str, "CnfFileParser: type of arg dir should be a list or a string"
            self.dirs = [dir]

    @property
    def files(self):
        for d in self.dirs:
            for f in sorted(listdir(d), reverse=True):
                if f.startswith("Production_"):
                    yield path.join(d, f)

    def get_cnf_file_rows(self, parts):
        prod_data = list()
        num_files = sum([len(listdir(x)) for x in self.dirs])
        with tqdm(desc="Fetching Data", total=num_files) as pbar:
            for processed_file in Pool().imap(self.file_worker, self.files):
                pbar.update()
                for line in processed_file:
                    part = part_name(*line[self.ipart.matl:self.ipart.job+1])
                    if part in parts:
                        line[self.ipart.matl] = part    # overwrite part with non-scan part name
                        prod_data.append(line)

        return prod_data

    def file_worker(self, f):
        result = list()
        with open(f, "r") as prod_file:
            for line in prod_file.readlines():
                result.append(line.upper().split("\t"))

        return result

    def parse_file(self, f):
        for row in self.file_worker(f):
            if f.startswith("Production_"):
                yield ParsedCnfRow(row)
            elif f.startswith("Issue_"):
                yield ParsedIssueRow(row)
            else:
                raise("Unmatched file type for parsing: {}".format(f))

    def parse_row(self, row):
        if type(row) is str:
            row = row.split("\t")

        return ParsedCnfRow(row, self.ipart, self.imatl)

class ParsedCnfRow:

    def __init__(self, row, part_indices=None, matl_indices=None):

        if part_indices and matl_indices:
            # part
            self.part_name = row[part_indices.matl]
            self.part_job = row[part_indices.job]
            self.part_wbs = row[part_indices.wbs]
            self.part_qty = int(row[part_indices.qty])

            # material
            self.matl_master = row[matl_indices.matl]
            self.matl_wbs = row[matl_indices.wbs]
            self.matl_qty = float(row[matl_indices.qty])
            self.matl_loc = row[matl_indices.loc]
            self.matl_plant = row[matl_indices.plant]

            self.program = row[part_indices.program]
        else: # from db
            # part
            self.part_name =      row[0]
            self.part_job =       row[1]
            self.part_wbs =       row[2]
            self.part_qty =   int(row[3])

            # material
            self.matl_master =    row[4]
            self.matl_wbs =       row[5]
            self.matl_qty = float(row[6])
            self.matl_loc =       row[7]
            self.matl_plant =     row[8]

            self.program =        row[9]

    @property
    def matl_qty_per_ea(self):
        return self.matl_qty / self.part_qty

    def __repr__(self):
        return "<ParsedCnfRow> {}".format(self.ls())

    def ls(self, qty=None, wbs=None):
        if len(wbs) == 5:
            wbs = "D-{}-{}".format(self.job[2:], wbs)

        return [
            #part
            self.part_name,
            self.part_job,
            wbs or self.part_wbs,
            "PROD",
            str(qty or self.part_qty),
            "EA",
            
            # material
            self.matl_master,
            self.matl_wbs,
            "{:.3f}".format( self.matl_qty_per_ea * (qty or self.part_qty) ),
            "IN2",
            self.matl_loc,
            self.matl_plant,
            
            self.program, "\n"
        ]

    def db_input(self):
        return [
            #part
            self.part_name,
            self.part_job,
            self.part_wbs,
            self.part_qty,
            
            # material
            self.matl_master,
            self.matl_wbs,
            self.matl_qty,
            self.matl_loc,
            self.matl_plant,
            
            self.program,
        ]

    def output(self, qty=None):
        result = "\t".join(self.ls(qty))

        self.matl_qty -= self.matl_qty_per_ea * (qty or self.part_qty)
        self.part_qty -= (qty or self.part_qty)

        return result

    def __add__(self, other):
        assert self.part_name == other.part_name, "Cannot add parts with different Marks"
        assert self.matl_master == other.matl_master, "Cannot add parts with different Material Masters"
        assert self.matl_wbs == other.matl_wbs, "Cannot add parts with different WBS Elements"

        self.part_qty += other.part_qty
        self.matl_qty += other.matl_qty


class ParsedIssueRow:

    def __init__(self, row, part_indices=None, matl_indices=None):
        index = SimpleNamespace(code=0, job=1, ship=2, program=9, matl=3, qty=5, loc=7, wbs=4, plant=8)

        if part_indices and matl_indices:
            # part
            self.code =      row[index.code]
            self.user1 =     row[index.job]
            self.user2 =     row[index.ship]

            # material
            self.matl =      row[index.matl]
            self.wbs =       row[index.wbs]
            self.qty = float(row[index.qty])
            self.loc =       row[index.loc]
            self.plant =     row[index.plant]

            self.program =   row[index.program]
        else: # from db
            # part
            self.code =      row[0]
            self.user1 =     row[1]
            self.user2 =     row[2]

            # material
            self.matl =      row[3]
            self.wbs =       row[4]
            self.qty = float(row[5])
            self.loc =       row[7]
            self.plant =     row[8]

            self.program =   row[9]

    def __repr__(self):
        return "<ParsedCnfRow> {}".format(self.ls())

    def ls(self, qty=None, wbs=None):
        if len(wbs) == 5:
            wbs = "D-{}-{}".format(self.job[2:], wbs)

        return [
            #part
            self.code,
            self.user1,
            self.user2,
            
            # material
            self.matl,
            self.wbs,
            "{:.3f}".format( self.qty ),
            "IN2",
            self.loc,
            self.plant,
            
            self.program, "\n"
        ]

    def db_input(self):
        return [
            #part
            self.code,
            self.user1,
            self.user2,
            
            # material
            self.matl,
            self.wbs,
            self.qty,
            self.loc,
            self.plant,
            
            self.program,
        ]

