import json
import urllib
import datetime
import openpyxl

TEMPLATE_FILE = 'template.xlsx'
URL = "https://api.coinmarketcap.com/v1/ticker/"

COLUMN_INDEX = 'A'
COLUMN_NAME = 'B'
COLUMN_SYMBOL = 'C'
COLUMN_PRICEUSD = 'D'
COLUMN_PRICEBTC = 'E'
COLUMN_MARKETCAPUSD = 'F'
COLUMN_24HVOLUMEUSD = 'G'
COLUMN_PERCCHANGE1H = 'H'
COLUMN_PERCCHANGE24H = 'I'
COLUMN_PERCCHANGE7D = 'J'
COLUMN_MAXSUPPLY = 'K'
COLUMN_TOTALSUPPLY = 'L'
COLUMN_AVAILABLESUPPLY = 'M'
COLUMN_RANK = 'N'
COLUMN_LASTUPDATED = 'O'

def storeData(sheet, row, col, data):
	if data is None: return
	sheet[col+str(row)] = data

response = urllib.urlopen(URL)
data = json.loads(response.read())

xfile = openpyxl.load_workbook(TEMPLATE_FILE)

sheet_stats = xfile.get_sheet_by_name('Statistics')
sheet_stats['B2'] = datetime.datetime.now()
sheet_stats['B3'] = URL
sheet = xfile.get_sheet_by_name('Data')
row = 2
for x in data:
	print x
	storeData(sheet, row, COLUMN_INDEX, row-1)
	storeData(sheet, row, COLUMN_NAME, x['name'])
	storeData(sheet, row, COLUMN_SYMBOL, x['symbol'])
	storeData(sheet, row, COLUMN_PRICEUSD, x['price_usd'])
	storeData(sheet, row, COLUMN_PRICEBTC, x['price_btc'])
	storeData(sheet, row, COLUMN_MARKETCAPUSD, x['market_cap_usd'])
	storeData(sheet, row, COLUMN_24HVOLUMEUSD, x['24h_volume_usd'])
	storeData(sheet, row, COLUMN_PERCCHANGE1H, x['percent_change_1h'])
	storeData(sheet, row, COLUMN_PERCCHANGE24H, x['percent_change_24h'])
	storeData(sheet, row, COLUMN_PERCCHANGE7D, x['percent_change_7d'])
	storeData(sheet, row, COLUMN_MAXSUPPLY, x['max_supply'])
	storeData(sheet, row, COLUMN_TOTALSUPPLY, x['total_supply'])
	storeData(sheet, row, COLUMN_AVAILABLESUPPLY, x['available_supply'])
	storeData(sheet, row, COLUMN_RANK, x['rank'])
	storeData(sheet, row, COLUMN_LASTUPDATED, x['last_updated'])
	row += 1
xfile.save('output.xlsx')

exit()

rb = open_workbook(TEMPLATE_FILE)
#wb = copy(rb)
wb = rb

s = wb.get_sheet(0)
row = 1
for x in data:
	print x
	s.write(row, COLUMN_INDEX, row)
	s.write(row, COLUMN_NAME, x['name'])
	s.write(row, COLUMN_SYMBOL, x['symbol'])
	s.write(row, COLUMN_PRICEUSD, x['price_usd'])
	s.write(row, COLUMN_PRICEBTC, x['price_btc'])
	row += 1
wb.save('output.xls')
