#!/usr/bin/env python

# cryptocurrencies2excel
# 
# Author: Enrico Cambiaso
# Email: enrico.cambiaso[at]gmail.com
# GitHub project URL: https://github.com/auino/cryptocurrencies2excel
# 

import json
import datetime
import openpyxl
import requests

###########################
### CONFIGURATION BEGIN ###
###########################

# conversion fiat currency
CURRENCY = 'USD'

# debug mode enabled?
DEBUG = False

# ho many pages to scroll? 0 to match all symbols in wallet
MAX_PAGES = 0

###########################
### CONFIGURATION END ###
###########################

TEMPLATE_FILE = 'template.xlsx'
WALLET_FILE = 'wallets.json'

BASE_URL = "https://coinmarketcap.com/?page={}"

COLUMN_INDEX = 1
COLUMN_NAME = 2
COLUMN_SYMBOL = 3
COLUMN_PRICE = 4
COLUMN_MARKETCAP = 5
COLUMN_24HVOLUME = 6
COLUMN_PERCCHANGE1H = 7
COLUMN_PERCCHANGE24H = 8
COLUMN_PERCCHANGE7D = 9
COLUMN_MAXSUPPLY = 10
COLUMN_TOTALSUPPLY = 11
COLUMN_CIRCULATINGSUPPLY = 12
COLUMN_RANK = 13
COLUMN_LASTUPDATED = 14

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
	sheet.cell(row=row, column=col).value = d

def transformtodata(data):
	r = []
	# getting keys
	data = data.get('props').get('dehydratedState').get('queries')[1].get('state').get('data').get('data').get('listing').get('cryptoCurrencyList')
	r = {}
	for e in data:
		k = e.get('symbol')
		r[k] = {'symbol':k, 'name':e.get('name'), 'circulatingSupply':e.get('circulatingSupply'), 'rank':e.get('cmcRank'), 'maxSupply':e.get('maxSupply'), 'totalSupply':e.get('totalSupply'), 'lastUpdated':e.get('lastUpdated')}
		if DEBUG: r[k]['data'] = e
		for q in e.get('quotes'):
			r[k]['quote.{}.price'.format(q.get('name'))] = q.get('price')
			r[k]['quote.{}.marketCap'.format(CURRENCY.upper())] = q.get('marketCap')
			r[k]['quote.{}.volume24h'.format(CURRENCY.upper())] = q.get('volume24h')
			r[k]['quote.{}.percentChange1h'.format(CURRENCY.upper())] = q.get('percentChange1h')
			r[k]['quote.{}.percentChange24h'.format(CURRENCY.upper())] = q.get('percentChange24h')
			r[k]['quote.{}.percentChange7d'.format(CURRENCY.upper())] = q.get('percentChange7d')
	return r

nowDateTime = datetime.datetime.now()
todayDate = nowDateTime.strftime("%m/%d/%Y")
todayTime = nowDateTime.strftime("%I:%M %p")

xfile = openpyxl.load_workbook(TEMPLATE_FILE)

walletdata = json.load(open(WALLET_FILE))
sheets_wallets = xfile['Wallets']
if len(walletdata) > 50:
	print('Truncating data to 50 elements')
	walletdata = walletdata[:50]

wallet_symbols = []
row = 4
for w in walletdata:
	wallet_symbols.append(w['symbol'])
	sheets_wallets['A'+str(row)] = w['symbol']
	sheets_wallets['B'+str(row)] = w['amount']
	try: sheets_wallets['C'+str(row)] = w['description']
	except: pass
	try: sheets_wallets['D'+str(row)] = ('' if w['reference'] is None else w['reference'])
	except: pass
	row += 1

import time
time.sleep(10)
sheets_wallets['B55'] = todayDate
sheets_wallets['C55'] = todayTime

sheet_stats = xfile['Statistics']
sheet_stats['B2'] = todayDate+' '+todayTime
sheet_stats['B3'].value = '=HYPERLINK("'+BASE_URL+'", "'+BASE_URL+'")'
sheet = xfile['Data']

stored_cryptos_count = 0

page = 0
while True:
	page += 1
	if MAX_PAGES != 0 and page > MAX_PAGES: break
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

	for k in data:
		row = stored_cryptos_count + 2
		r = data.get(k)
		if DEBUG: print(r)
		if MAX_PAGES == 0 and r['symbol'] in wallet_symbols:
			wallet_symbols.remove(r['symbol'])
			if len(wallet_symbols) == 0: break
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
		storeData(sheet, row, COLUMN_LASTUPDATED, r.get('lastUpdated'), tostr)
		stored_cryptos_count += 1
	else: continue
	break
xfile.save('output.xlsx')
