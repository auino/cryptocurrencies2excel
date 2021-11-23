#!/usr/bin/env python

# cryptocurrencies2excel
# 
# Author: Enrico Cambiaso
# Email: enrico.cambiaso[at]gmail.com
# GitHub project URL: https://github.com/auino/cryptocurrencies2excel
# 

import sys

PATH_LIST = [
	'/usr/local/Cellar/python/3.7.7/Frameworks/Python.framework/Versions/3.7/lib/python37.zip',
	'/usr/local/Cellar/python/3.7.7/Frameworks/Python.framework/Versions/3.7/lib/python3.7',
	'/usr/local/Cellar/python/3.7.7/Frameworks/Python.framework/Versions/3.7/lib/python3.7/lib-dynload',
	'/Users/enricocambiaso/Library/Python/3.7/lib/python/site-packages',
	'/usr/local/lib/python3.7/site-packages'
]

for e in PATH_LIST:
	if not e in sys.path: sys.path.append(e)

import json
import datetime
import openpyxl

try: import requests
except: import workflow.web as requests

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

# ho many pages to scroll?
MAX_PAGES = 2

###########################
### CONFIGURATION END ###
###########################

TEMPLATE_FILE = 'template.xlsx'
WALLET_FILE = 'wallets.json'

#BASE_URL = "https://api.coinmarketcap.com/v1/ticker/?convert="+CURRENCY.upper()
BASE_URL = "https://coinmarketcap.com/?page={}"

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

def transformtodata(data):
	r = []
	# getting keys
	keys = data['props']['initialState']['cryptocurrency']['listingLatest']['data'][0]['keysArr']
	keys += data['props']['initialState']['cryptocurrency']['listingLatest']['data'][0]['excludeProps']
	#for k in data['props']['initialState']['cryptocurrency']['listingLatest']['data'][0]['excludeProps']: keys.remove(k)
	values = data['props']['initialState']['cryptocurrency']['listingLatest']['data'][1:]
	for v in values:
		e = {}
		for i in range(0, len(v)): e[keys[i]] = v[i]
		r.append(e)
	return r

nowDateTime = datetime.datetime.now()
todayDate = nowDateTime.strftime("%m/%d/%Y")
todayTime = nowDateTime.strftime("%I:%M %p")

xfile = openpyxl.load_workbook(TEMPLATE_FILE)

walletdata = json.load(open(WALLET_FILE))
sheets_wallets = xfile['Wallets']
if len(walletdata) > 20: walletdata = walletdata[:20]
row = 4
for w in walletdata:
	sheets_wallets['A'+str(row)] = w['symbol']
	sheets_wallets['B'+str(row)] = w['amount']
	try: sheets_wallets['C'+str(row)] = w['description']
	except: pass
	row += 1

import time
time.sleep(10)
sheets_wallets['B25'] = todayDate
sheets_wallets['C25'] = todayTime

sheet_stats = xfile['Statistics']
sheet_stats['B2'] = todayDate+' '+todayTime
sheet_stats['B3'].value = '=HYPERLINK("'+BASE_URL+'", "'+BASE_URL+'")'
sheet = xfile['Data']

stored_cryptos_count = 0

for page in range(1, MAX_PAGES+1):
	URL = BASE_URL.format(page)
	response = requests.get(URL).text
	response = response.split('<script ')
	for r in response:
		if '"priceChange"' in r:
			r = r[r.index('>')+1:]
			r = r[:r.index('</script>')]
			response = r
			break
	if DEBUG: print(response)

	data = json.loads(response)
	data = transformtodata(data)

	for r in data:
		row = stored_cryptos_count + 2
		if DEBUG: print(r)
		storeData(sheet, row, COLUMN_INDEX, row-1, toint)
		storeData(sheet, row, COLUMN_NAME, r['name'], tostr)
		storeData(sheet, row, COLUMN_SYMBOL, r['symbol'], tostr)
		storeData(sheet, row, COLUMN_PRICE, r['quote.{}.price'.format(CURRENCY.upper())], tofloat)
		storeData(sheet, row, COLUMN_MARKETCAP, r['quote.{}.marketCap'.format(CURRENCY.upper())], tofloat)
		storeData(sheet, row, COLUMN_24HVOLUME, r['quote.{}.volume24h'.format(CURRENCY.upper())], tofloat)
		storeData(sheet, row, COLUMN_PERCCHANGE1H, r['quote.{}.percentChange1h'.format(CURRENCY.upper())], tofloat)
		storeData(sheet, row, COLUMN_PERCCHANGE24H, r['quote.{}.percentChange24h'.format(CURRENCY.upper())], tofloat)
		storeData(sheet, row, COLUMN_PERCCHANGE7D, r['quote.{}.percentChange7d'.format(CURRENCY.upper())], tofloat)
		storeData(sheet, row, COLUMN_MAXSUPPLY, r.get('maxSupply'), tofloat)
		storeData(sheet, row, COLUMN_TOTALSUPPLY, r.get('totalSupply'), tofloat)
		storeData(sheet, row, COLUMN_CIRCULATINGSUPPLY, r.get('circulatingSupply'), tofloat)
		storeData(sheet, row, COLUMN_RANK, r.get('rank'), toint)
		storeData(sheet, row, COLUMN_LASTUPDATED, r.get('lastUpdated'), toint)
		stored_cryptos_count += 1

xfile.save('output.xlsx')
