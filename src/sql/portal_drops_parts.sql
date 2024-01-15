select distinct
	ProgramName, PartName, QtyInProcess
from PIPArchive
where PartName like '1210027E%'
	and TransType = 'SN102'
	and WONumber like '1210027E-7-%'
