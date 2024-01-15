SELECT * FROM
(
    SELECT
        Program.ArcDateTime     AS Timestamp,
        Stock.Location          AS Location,
        Stock.PrimeCode         AS Matl,
        Stock.SheetName         AS Sheet,
        Program.ProgramName     AS Program,
        Stock.HeatNumber        AS Heat,
        Stock.BinNumber         AS PO
    FROM StockHistory AS Stock
        INNER JOIN ProgArchive AS Program
            ON Stock.SheetName=Program.SheetName
    WHERE
        Stock.SheetName LIKE ?
    AND
        Program.TransType='SN102'

    UNION

    SELECT
        0                       AS Timestamp,
        Stock.Location          AS Location,
        Stock.PrimeCode         AS Matl,
        Stock.SheetName         AS Sheet,
        Program.ProgramName     AS Program,
        Stock.HeatNumber        AS Heat,
        Stock.BinNumber         AS PO
    FROM Stock
        LEFT JOIN Program
            ON Stock.SheetName=Program.SheetName
    WHERE
        Stock.SheetName LIKE ?
) AS Stock
ORDER BY Timestamp DESC