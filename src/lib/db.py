
from typing import Any
import pyodbc

from os import getenv, path
from datetime import datetime
from string import Template

from lib.part import Part

SNDB_PRD = "HSSSNData"
SNDB_DEV = "HIIWINBL5"

CONN_STR_USER_AUTH = Template("DRIVER={$driver};SERVER=$server;UID=$user;PWD=$pwd;DATABASE=$db;")
CONN_STR_WIN_AUTH = Template("DRIVER={$driver};SERVER=$server;Trusted_connection=yes;")

if len(pyodbc.drivers()) > 0:
    DEFAULT_DRIVER = pyodbc.drivers()[0]
else:
    raise NotImplementedError("No SQL drivers available")

class DbConnection:
    """
        pyodbc connection wrapper for databases
    """

    def __init__(self, use_win_auth=False, **kwargs):
        if use_win_auth:
            self.CS_TEMP = CONN_STR_WIN_AUTH
        else:
            self.CS_TEMP = CONN_STR_USER_AUTH

        self.driver = DEFAULT_DRIVER
        self.__dict__.update(kwargs)

        self._cnxn = None
        self._cur = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self._cnxn is not None:
            self._cnxn.close()

    def _make_cnxn(self):
        self._cnxn = pyodbc.connect(self.CS_TEMP.substitute(**self.__dict__))

    def __getattr__(self, name):
        """
        any attribute (or function) not defined by this class
        will be called if it exists for the (in order)
            1) Cursor
            2) Connection
        """
        if hasattr(self.cursor, name):
            return getattr(self.cursor, name)
        
        if hasattr(self.connection, name):
            return getattr(self.connection, name)
        
        raise AttributeError

    @property
    def connection(self):
        if not self._cnxn:
            self._make_cnxn()

        return self._cnxn

    @property
    def cursor(self):
        if not self._cur:
            self._cur = self.connection.cursor()

        return self._cur

    def execute_sql_file(self, file_path, *args):
        with open(file_path, 'r') as sql_file:
            sql = sql_file.read()

            self.cursor.execute(sql, *args)

        return self
    
    def query_from_sql_file(self, file_path, *args):
        assert path.exists(file_path), "SQL file `{}` does not exist".format(file_path)

        with open(file_path, 'r') as sql_file:
            sql = sql_file.read()

            return self.cursor.execute(sql, *args)

    def collect_table_data(self):
        min_date = datetime(1900, 1, 1)

        data = [None]
        for row in self.cursor.fetchall():
            data.append(row)

            # replace active dates (1900-01-01 00:00:00 -> --)
            if min_date in data[-1]:
                data[-1] = ["--" if x == min_date else x for x in data[-1]]
        else:
            try:
                data[0] = [t[0] for t in row.cursor_description]
            except UnboundLocalError:
                caller = self.__dict__.get("func", "-- no origin --")
                return [["values"], ["nothing returned ({})".format(caller)]]

        return data

class SndbConnection(DbConnection):
    """
        db connection wrapper for SigmaNest databases
    """

    def __init__(self, dev=False, **kwargs):
        init_kwargs = dict(
            server=SNDB_PRD,
            db="SNDBase91",
            user=getenv('SNDB_USER'),
            pwd=getenv('SNDB_PWD'),
        )

        init_kwargs.update(kwargs)
        super().__init__(**init_kwargs)

        if dev:
            self.server = SNDB_DEV
            self.db = "SNDBaseDev"

class BomConnection(DbConnection):
    """
        db connection wrapper for the engineering BOM databse
    """

    def __init__(self, **kwargs):
        init_kwargs = dict(
            server="HSSSQLSERV",
        )

        init_kwargs.update(kwargs)
        super().__init__(use_win_auth=True, **init_kwargs)

    def get_bom(self, job, shipment, mark=None):
        self.cursor.execute(
            "EXEC BOM.SAP.GetBOMData @Job=?, @Ship=?",
            job, shipment
        )

        if mark:
            index = None
            mark = mark.casefold()
            for row in self.cursor.fetchall():
                index = index or [t[0] for t in row.cursor_description].index("Piecemark")
                
                if row[index].casefold() == mark:
                    return Part(row)
            else:
                return None
        
        # else
        return [Part(row) for row in self.cursor.fetchall()]


def bom(job, shipment, mark=None):
    BomConnection().get_bom(job, shipment, mark)
