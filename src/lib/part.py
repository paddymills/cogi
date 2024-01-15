
from fractions import Fraction
from pyodbc import Row
from re import compile as regex

aliases = dict(
    mark    = ("Piecemark",),
    qty     = ("Qty",),
    desc    = ("Description",),
    type    = ("Commodity",),
    thk     = ("Thick",),
    wid     = ("Width",),
    len     = ("Length",),
    spec    = ("Specification",),
    grade   = ("Grade",),
    test    = ("ImpactTest",),
    remark  = ("Remark",),
    item    = ("Item",),
    dwg     = ("DwgNo",),
)

STOCK_GRADES = regex(r"^\w+-(50|345)W?(T|T2)?$")
STOCK_THK = [
    0.25,
    0.375,
    0.5,
    0.625,
    0.75,
    0.875,
    1.0,
    1.125,
    1.25,
    1.5,
    2.0,
]
MILL_THK_BUFFER = 0.125

def float_display(f, display_feet=False, force_zero=False):
    if display_feet and f >= 12.0:
        return "{:g}'-{}".format(f // 12, float_display(f % 12, force_zero=True))

    whole = int(f)
    frac = Fraction(f % 1)

    if whole == 0 and not force_zero:
        whole = ''

    if frac == 0:
        frac = ''

    return "{} {}".format(whole, frac).strip()


class Part:

    def __init__(self, init_data=None):
        if init_data is not None:
            self.parse_data(init_data)

    @property
    def matl_grade(self):
        if self.spec == 'A240 Type 304':
            self.spec = 'A240'
            self.grade = '304'
        if self.spec == 'A240 Type 316':
            self.spec = 'A240'
            self.grade = '316'
        if self.spec == 'A606 TYPE 4':
            self.spec = 'A606'
            self.grade = 'TYPE4'

        if "SRM" in self.remark:
            self.test = "FCM"

        if not self.test:
            zone = ''
        elif "HPS" in self.grade:
            zone = '3'
        else:
            zone = '2'

        if self.test == 'FCM':
            self.test = 'F'


        return "{}-{}{}{}".format(self.spec, self.grade, self.test, zone)

    @property
    def matl_grade_cvn(self):
        if self.test or self.spec in ('A240', 'A606'):
            return self.matl_grade
        
        if "HPS" in self.grade:
            return self.matl_grade + "T3"

        return self.matl_grade + "T2"

    def for_prenest(self, prenest_type='all', skip_stock=True):
        # exclude
        #   - secondary from main prenest
        #   - main from secondary prenest
        if self.is_main:
            if prenest_type == 'secondary':
                return False
        else:  # is secondary
            if prenest_type == 'main':
                return False

        # exclude shapes
        if self.type not in ("PL", "SHT"):
            return False

        if skip_stock:
            return True

        # only place stock plate match tests below here
        # ---------------------------------------------

        # not compatible with 50/50WT2
        if not STOCK_GRADES.match(self.matl_grade):
            return True

        # larger than stock plate size
        if self.wid > 95.0:
            return True
        if self.thk <= 0.375 and self.len > 120.0:
            return True
        elif self.len > 240.0:
            return True

        # if part gets thickness mill and
        #   is within MILL_THK_BUFFER of a stock thickness
        if self.gets_thk_mill:
            for thk in STOCK_THK:
                if 0 < thk - self.thk <= MILL_THK_BUFFER:
                    return False
        elif self.thk in STOCK_THK:
            return False

        return True

    @property
    def is_main(self):
        return self.item.startswith('03') or self.item.startswith('04')

    @property
    def gets_thk_mill(self):
        if self.remark.startswith("BEV. FILL"):
            return True

        # TODO: verify this one
        # if "M1F" in self.remark:
        #     return True

        return False

    def __repr__(self):
        return "Part<{}: {} [{}] [{}]>".format(self.mark, self.size, self.matl_grade, self.item)

    @property
    def size(self):
        _thk = float_display(self.thk)
        _wid = float_display(self.wid)
        _len = float_display(self.len, display_feet=True)

        return "{} {} x {} x {}".format(self.type, _thk, _wid, _len)

    def xml_format(self):
        return (self.mark, self.qty, self.thk, self.wid, self.len, self.matl_grade_cvn,
                self.item, self.dwg, *[None] * 5, self.remark, float_display(self.len * self.qty, display_feet=True))

    def parse_data(self, data):
        if type(data) is dict:
            self.__dict__.update(data)
        
        elif type(data) is Row:
            self._parse_row(data)

        else:
            raise NotImplementedError("unmatched data type")

    def _parse_row(self, row):
        header = [t[0] for t in row.cursor_description]

        for k, v in aliases.items():
            index = self.get_index_by_alias(header, v)
            setattr(self, k, row[index])

    def get_index_by_alias(self, row, aliases):
        for a in aliases:
            try:
                return row.index(a)
            except ValueError:
                pass

        raise IndexError("No index found for aliases: {}".format(aliases))
