#!/usr/bin/env python

# cryptocurrencies2excel
# 
# Author: Enrico Cambiaso
# Email: enrico.cambiaso[at]gmail.com
# GitHub project URL: https://github.com/auino/cryptocurrencies2excel
# 

import json
import requests
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
DEBUG = True

###########################
### CONFIGURATION END ###
###########################

TEMPLATE_FILE = 'template.xlsx'
WALLET_FILE = 'wallets.json'

#BASE_URL = "https://api.coinmarketcap.com/v1/ticker/?convert="+CURRENCY.upper()
BASE_URL = "https://coinmarketcap.com"

COLUMN_INDEX = 'A'
COLUMN_NAME = 'B'
COLUMN_SYMBOL = 'C'
COLUMN_PRICE = 'D'
COLUMN_MARKETCAP = 'E'
COLUMN_24HVOLUME = 'F'
COLUMN_PERCCHANGE1H = 'G'
COLUMN_PERCCHANGE24H = 'H'
COLUMN_PERCCHANGE7D = 'I'
COLUMN_MAXSUPPLY = 'J'
COLUMN_TOTALSUPPLY = 'K'
COLUMN_CIRCULATINGSUPPLY = 'L'
COLUMN_RANK = 'M'
COLUMN_LASTUPDATED = 'N'

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

URL = BASE_URL
response = requests.get(URL).text
response = response.split('<script ')
for r in response:
	if "__NEXT_DATA__" in r:
		response = r
		response = response.split('>')[1]
		response = response.split('</script')[0]
		break
if DEBUG: print(response)

data = json.loads(response)

stored_cryptos_count = 0
for r in data['props']['initialState']['cryptocurrency']['listingLatest']['data']:
	row = stored_cryptos_count + 2
	if DEBUG: print(r)
	storeData(sheet, row, COLUMN_INDEX, row-1, toint)
	storeData(sheet, row, COLUMN_NAME, r['name'], tostr)
	storeData(sheet, row, COLUMN_SYMBOL, r['symbol'], tostr)
	storeData(sheet, row, COLUMN_PRICE, r['quote'][CURRENCY.upper()]['price'], tofloat)
	storeData(sheet, row, COLUMN_MARKETCAP, r['quote'][CURRENCY.upper()]['market_cap'], tofloat)
	storeData(sheet, row, COLUMN_24HVOLUME, r['quote'][CURRENCY.upper()]['volume_24h'], tofloat)
	storeData(sheet, row, COLUMN_PERCCHANGE1H, r['quote'][CURRENCY.upper()]['percent_change_1h'], tofloat)
	storeData(sheet, row, COLUMN_PERCCHANGE24H, r['quote'][CURRENCY.upper()]['percent_change_24h'], tofloat)
	storeData(sheet, row, COLUMN_PERCCHANGE7D, r['quote'][CURRENCY.upper()]['percent_change_7d'], tofloat)
	storeData(sheet, row, COLUMN_MAXSUPPLY, r['max_supply'], tofloat)
	storeData(sheet, row, COLUMN_TOTALSUPPLY, r['total_supply'], tofloat)
	storeData(sheet, row, COLUMN_CIRCULATINGSUPPLY, x['circulating_supply'], tofloat)
	storeData(sheet, row, COLUMN_RANK, r['rank'], toint)
	storeData(sheet, row, COLUMN_LASTUPDATED, r['last_updated'], toint)
	stored_cryptos_count += 1

xfile.save('output.xlsx')
