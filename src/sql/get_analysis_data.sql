DECLARE @PrevWeekSunday DATETIME
DECLARE @ThisWeekSunday DATETIME
SELECT @PrevWeekSunday = DATEADD(wk, DATEDIFF(wk, 6, GETDATE()), -1)
SELECT @ThisWeekSunday = DATEADD(wk, DATEDIFF(wk, 0, GETDATE()), -1)

SELECT
	part.AutoID AS Id,
	program.ArcDateTime AS UpdateDate,
    REPLACE(PartName, '_', '-') AS Part,
    part.ProgramName AS Program,
    QtyProgram AS Qty,
    ROUND(NestedArea * QtyProgram, 3) AS Area,

    stock.Location,
    stock.PrimeCode AS MaterialMaster,
    stock.Mill AS Wbs,
    
    CASE LEFT(program.MachineName,7)
        WHEN 'Plant_3' THEN 'HS02'
        ELSE 'HS01'
    END AS Plant,
	'' AS OrderOrDocument,
	'' AS SAPValue,
	'' AS Notes
FROM PartArchive AS part
    INNER JOIN StockArchive AS stock
        ON part.ArchivePacketID=stock.ArchivePacketID
    INNER JOIN ProgArchive AS program
        ON part.ArchivePacketID=program.ArchivePacketID
        AND program.TransType='SN102'
WHERE program.ArcDateTime >= @PrevWeekSunday
AND program.ArcDateTime < @ThisWeekSunday
ORDER BY stock.PrimeCode, part.ProgramName, PartName