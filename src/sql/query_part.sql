SELECT
    PIP.ArcDateTime     AS Timestamp,
    PIP.PartName        AS Part,
    PIP.QtyInProcess    AS Qty,
    PIP.ProgramName     AS Program,
    Stock.HeatNumber    AS Heat,
    Stock.BinNumber     AS PO,
    Stock.PrimeCode     AS Matl
FROM PIPArchive AS PIP
    INNER JOIN StockArchive AS Stock
        ON PIP.ProgramName=Stock.ProgramName
WHERE
    PIP.PartName LIKE ?
AND
    PIP.TransType='SN102'

UNION

SELECT
    0                   AS Timestamp,
    PIP.PartName        AS Part,
    PIP.QtyInProcess    AS Qty,
    PIP.ProgramName     AS Program,
    Stock.HeatNumber    AS Heat,
    Stock.BinNumber     AS PO,
    Stock.PrimeCode     AS Matl
FROM PIP
    INNER JOIN Program
        ON PIP.ProgramName=Program.ProgramName
    INNER JOIN Stock
        ON Program.SheetName=Stock.SheetName
WHERE
    PIP.PartName LIKE ?
ORDER BY PIP.PartName, PIP.ArcDateTime
