select
	case
		when num_jobs > 1 then '<MULTIPLE JOBS>'
		when stock.PrimeCode like '%-03%' then 'Main'
		when stock.PrimeCode like '%-04%' then 'Main'
		else 'Secondary'
	end as MaterialGroup,

	stock.PrimeCode as MaterialMaster,
	prog.ProgramName,
	null as PartName,	-- to be calculated in script
	null as PartQty,	-- to be calculated in script

	case
		when stock.SheetType = 0 then 'New'
		when stock.SheetType = 1 then 'Remnant'
		else '<undefined>'
	end as StockType,

	stock.Thickness as SheetThickness,
	stock.Width as SheetWidth,
	stock.Length as SheetLength,
	stock.Area as SheetArea,

	prog.UsedArea,
	prog.ScrapFraction,
	null as ScrapArea,		-- to be calculated in script
	null as ScrapWeight		-- to be calculated in script
	-- prog.ScrapFraction * prog.UsedArea as ScrapArea,
	-- prog.ScrapFraction * prog.UsedArea * stock.Thickness * 0.2836 as ScrapWeight

from ProgArchive as prog
	inner join StockArchive as stock
		on stock.ProgramName=prog.ProgramName
		and stock.SheetName=prog.SheetName
	inner join (
		select
			ProgramName,
			count(job) as num_jobs
		from (
			select distinct
				pip.ProgramName,
				p.Data1 as job
			from PIPArchive as pip
				inner join Part as p
					on pip.PartName=p.PartName
			where
				pip.WONumber like '1210027E-7-%'
				and pip.TransType = 'SN102'
		) as a
		group by ProgramName
	) as job_count
		on job_count.ProgramName=prog.ProgramName
where prog.TransType = 'SN102'
order by MaterialMaster, ProgramName