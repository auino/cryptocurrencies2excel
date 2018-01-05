#!/usr/bin/env python

# cryptocurrencies2excel
# 
# Author: Enrico Cambiaso
# Email: enrico.cambiaso[at]gmail.com
# GitHub project URL: https://github.com/auino/cryptocurrencies2excel
# 

import json
import urllib
import datetime
import openpyxl

###########################
### CONFIGURATION BEGIN ###
###########################

# conversion fiat currency
CURRENCY = 'USD'

# how many cryptos to be retrieved?
CRYPTOS_NUMBER = 1000
# how many crypto data for each request?
CRYPTOS_LIMIT = 100

# debug mode enabled?
DEBUG = False

###########################
### CONFIGURATION END ###
###########################

TEMPLATE_FILE = 'template.xlsx'
WALLET_FILE = 'wallets.json'

BASE_URL = "https://api.coinmarketcap.com/v1/ticker/?convert="+CURRENCY.upper()

COLUMN_INDEX = 'A'
COLUMN_NAME = 'B'
COLUMN_SYMBOL = 'C'
COLUMN_PRICE = 'D'
COLUMN_PRICEBTC = 'E'
COLUMN_MARKETCAP = 'F'
COLUMN_24HVOLUME = 'G'
COLUMN_PERCCHANGE1H = 'H'
COLUMN_PERCCHANGE24H = 'I'
COLUMN_PERCCHANGE7D = 'J'
COLUMN_MAXSUPPLY = 'K'
COLUMN_TOTALSUPPLY = 'L'
COLUMN_AVAILABLESUPPLY = 'M'
COLUMN_RANK = 'N'
COLUMN_LASTUPDATED = 'O'

def toint(v):
	try: return int(v)
	except: return None
def tofloat(v):
	try: return float(v)
	except: return None
def tostr(v):
	try: return str(v)
	except: return None

def storeData(sheet, row, col, data, f):
	d = f(data)
	if d is None: return
	sheet[col+str(row)] = d

nowDateTime = datetime.datetime.now()
todayDate = nowDateTime.strftime("%m/%d/%Y")
todayTime = nowDateTime.strftime("%I:%M %p")

xfile = openpyxl.load_workbook(TEMPLATE_FILE)

walletdata = json.load(open(WALLET_FILE))
sheets_wallets = xfile.get_sheet_by_name('Wallets')
if len(walletdata) > 10: walletdata = walletdata[:10]
row = 4
for w in walletdata:
	sheets_wallets['A'+str(row)] = w['symbol']
	sheets_wallets['B'+str(row)] = w['amount']
	try: sheets_wallets['C'+str(row)] = w['description']
	except: pass
	row += 1
sheets_wallets['B15'] = todayDate
sheets_wallets['C15'] = todayTime

sheet_stats = xfile.get_sheet_by_name('Statistics')
sheet_stats['B2'] = todayDate+' '+todayTime
sheet_stats['B3'].value = '=HYPERLINK("'+BASE_URL+'", "'+BASE_URL+'")'
sheet = xfile.get_sheet_by_name('Data')
stored_cryptos_count = 0
while stored_cryptos_count < CRYPTOS_NUMBER:
	URL = BASE_URL+"&start="+str(stored_cryptos_count)+"&limit="+str(CRYPTOS_LIMIT)
	if DEBUG: print URL
	response = urllib.urlopen(URL)
	data = json.loads(response.read())
	for x in data:
		if DEBUG: print x
		row = stored_cryptos_count + 2
		storeData(sheet, row, COLUMN_INDEX, row-1, toint)
		storeData(sheet, row, COLUMN_NAME, x['name'], tostr)
		storeData(sheet, row, COLUMN_SYMBOL, x['symbol'], tostr)
		storeData(sheet, row, COLUMN_PRICE, x['price_'+CURRENCY.lower()], tofloat)
		storeData(sheet, row, COLUMN_PRICEBTC, x['price_btc'], tofloat)
		storeData(sheet, row, COLUMN_MARKETCAP, x['market_cap_'+CURRENCY.lower()], tofloat)
		storeData(sheet, row, COLUMN_24HVOLUME, x['24h_volume_'+CURRENCY.lower()], tofloat)
		storeData(sheet, row, COLUMN_PERCCHANGE1H, x['percent_change_1h'], tofloat)
		storeData(sheet, row, COLUMN_PERCCHANGE24H, x['percent_change_24h'], tofloat)
		storeData(sheet, row, COLUMN_PERCCHANGE7D, x['percent_change_7d'], tofloat)
		storeData(sheet, row, COLUMN_MAXSUPPLY, x['max_supply'], tofloat)
		storeData(sheet, row, COLUMN_TOTALSUPPLY, x['total_supply'], tofloat)
		storeData(sheet, row, COLUMN_AVAILABLESUPPLY, x['available_supply'], tofloat)
		storeData(sheet, row, COLUMN_RANK, x['rank'], toint)
		storeData(sheet, row, COLUMN_LASTUPDATED, x['last_updated'], toint)
		stored_cryptos_count += 1
xfile.save('output.xlsx')

exit()

rb = open_workbook(TEMPLATE_FILE)
#wb = copy(rb)
wb = rb

s = wb.get_sheet(0)
row = 1
for x in data:
	if DEBUG: print x
	s.write(row, COLUMN_INDEX, row)
	s.write(row, COLUMN_NAME, x['name'])
	s.write(row, COLUMN_SYMBOL, x['symbol'])
	s.write(row, COLUMN_PRICE, x['price_'+CURRENCY.lower()])
	s.write(row, COLUMN_PRICEBTC, x['price_btc'])
	row += 1
wb.save('output.xls')
