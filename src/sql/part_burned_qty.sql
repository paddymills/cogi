SELECT
    WONumber, QtyInProcess
FROM PIPArchive
WHERE
    PartName LIKE ?
AND
    TransType = 'SN102'
AND
    ArcDateTime < ?