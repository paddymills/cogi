
from argparse import ArgumentParser
from datetime import date
from os import path
import pandas
import xlwings

from lib import db

onedrive = r"C:\Users\pmiller1\OneDrive - high.net\inventory_recon"
xl_file = r"C:\users\pmiller1\Documents\SAP\SAP GUI\mb52.XLSX"
query = """
DECLARE @LastIntervalHourTimestamp DATETIME
DECLARE @intervalHours INT
SELECT @intervalHours = 4
SELECT @LastIntervalHourTimestamp = DATEADD(HOUR, DATEDIFF(HOUR, 0, GETDATE()) / @intervalHours * @intervalHours, 0)


SELECT
	Plant, Location, MaterialMaster, Wbs,
	ROUND(SUM(Area), 3) AS Area
FROM (
	SELECT
		IIF(LEFT(SheetName, 1)='W', 'HS02', 'HS01') AS Plant,
		Location,
		PrimeCode AS MaterialMaster,
		Mill AS Wbs,
		Area
	FROM (
		SELECT SheetName, Location, PrimeCode, Mill, Qty * Area as Area FROM Stock
		UNION
		SELECT SheetName, Location, PrimeCode, Mill, Area FROM StockArchive
			WHERE ArcDateTime > @LastIntervalHourTimestamp
	) AS Stock
	WHERE
		SheetName NOT LIKE 'SLAB%'
) AS Stock
GROUP BY
	Plant, Location, MaterialMaster, Wbs
ORDER BY
	Plant, Location, MaterialMaster, Wbs
"""

sap_col_rename = {
	"Material Number": "MaterialMaster",
	"Storage Location": "Location",
	"Unrestricted": "Area",
	"Base Unit of Measure": "UoM",
	"Special stock number": "Wbs",
	"Plant": "Plant", 	# so that we have the key to keep this column
}
uom = ['IN2', 'FT2', 'M2']

def get_sn(export=False, from_csv=False):
	if from_csv:
		df = pandas.read_csv(r"C:\temp\rollup_sn.csv")
	else:
		data = dict()
		with db.SndbConnection() as conn:
			conn.cursor.execute(query)

			results = conn.collect_table_data()
			head = results.pop(0)
			for i, col in enumerate(head):
				data[col] = [r[i] for r in results]

		df = pandas.DataFrame(data)

		if export:
			df.to_csv(r"C:\temp\rollup_sn.csv", index=False)


	df = df.dropna(subset='MaterialMaster')
	df = df.set_index(["MaterialMaster", "Plant", "Location", "Wbs"])

	
	return df

def get_sap():
	df = pandas.read_excel(xl_file)
	df = df.rename(columns=sap_col_rename)
	df = df.filter(items=sap_col_rename.values())
	df = df.loc[df['UoM'].isin(uom)]
	df = df.loc[df['Location'] != "RAIL"]
	df = df.dropna(subset='MaterialMaster')
	df = df.set_index(["MaterialMaster", "Plant", "Location", "Wbs"])
	
	return df


def do_compare(sn, sap):
	sap = sap.rename(columns={"Area": "SAP"})
	sn = sn.rename(columns={"Area": "Sigmanest"})

	df = sap.join(sn,
		how='outer',
		lsuffix="_sap",
		rsuffix="_sn",
	)

	df = df.fillna(value={"UoM": "IN2"})

	df = df.reset_index(level='Plant')  # remove Plant from index
	df = df.sort_index()

	hs01 = df.loc[df['Plant'] == "HS01"]
	hs02 = df.loc[df['Plant'] == "HS02"]

	return hs01.drop(columns="Plant"), hs02.drop(columns="Plant")


def main():
	parser = ArgumentParser()
	parser.add_argument("-e", "--export", action="store_true", help="save Sigmanest state as csv file")
	parser.add_argument("-c", "--csv", action="store_true", help="read Sigmanest state from csv file")
	parser.add_argument("-o", "--overwrite", action="store_true", help="overwrite output file instead of appending count")
	parser.add_argument("-v", "--verbose", action="count", default=0, help="Verbosity level")
	args = parser.parse_args()

	sn = get_sn(export=args.export, from_csv=args.csv)
	if args.export:
		return

	sap = get_sap()
	lanc, wilpo = do_compare(sn, sap)

	if args.verbose > 0:
		print(sn)
		print(sap)
	if args.verbose > 1:
		print(lanc)
		print(wilpo)

	wb = xlwings.Book(path.join(onedrive, "template.xlsx"))

	datasets = [
		("Sigmanest", sn),
		("SAP", sap),
		("Lancaster", lanc),
		("Williamsport", wilpo),
	]
	for name, data in datasets:
		sheet = wb.sheets(name)
		sheet.range("A1").value = data
		sheet.autofit()

	save_name = date.today().strftime("%Y-%m-%d") + ".xlsx"
	if not args.overwrite:
		count = 1
		while path.exists(path.join(onedrive, save_name)):
			save_name = "{}_{}.xlsx".format(save_name[:10], count)
			count += 1

	print("save:", save_name)
	wb.save(path.join(onedrive, save_name))

	with open("temp/mm.txt", 'w') as f:
		f.write("\n".join(set([x[0] for x in lanc[1:]])))


if __name__ == "__main__":
	main()
