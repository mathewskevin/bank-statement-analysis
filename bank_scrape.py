# Kevin Mathews 8/16/2019 rev 1.02
# Bank Analysis Web Scraper
# written in Python 3

# Automated web scraper for gathering bank statement CSVs in 90 day chunks.

import webbrowser
from selenium import webdriver
import pdb, time, sys, os
import random, datetime
from datetime import date, timedelta
import pandas as pd

#config_path = "D:\My_Storage\Workspace\bank-account-analysis-master\bank_files"
config_path = "D:\My_Storage\My_Documents\bank-account-analysis-master\bank_files"
sys.path.append(config_path) 
import config_pass

# enter password
user_name = config_pass.BANK_USERNAME
user_pass = config_pass.BANK_PASSWORD

# input variables
site_name = 'https://www.navyfederal.org/' # site name
profile_loc = r'C:\Users\Kevin\AppData\Roaming\Mozilla\Firefox\Profiles\r5s0g7v9.default' # web browser profile
credit_name = 'Visa Signature GO REWARDS - 5295' # credit account name
checking_name = 'EveryDay Checking - 3969' # checking account name
savings_name = 'Share Savings - 8327' # savings account name

download_folder = r'C:\Users\Kevin\Downloads'
#database_folder = r'D:\My_Storage\Workspace\navyfed-account-analysis-master'
database_folder = r"D:\My_Storage\My_Documents\bank-account-analysis-master"
database_file = database_folder + '\\database_file.xlsx' # database_file (please enable editing)
src = download_folder + '\\transactions.CSV'

# calculate random time to simulate human input
def rand_time(lower, upper):
	return random.uniform(lower, upper)

# function to add date sorting column to data
def add_date(row):
	if data_type == 'credit':
		date_val = row['Transaction Date']
	else:
		date_val = row['Date']
	
	date_format = datetime.datetime.strptime(date_val, '%m/%d/%Y')
	date_month = date_format.month
	if date_month < 10:
		date_month = '0' + str(date_month)
	else:
		date_month = str(date_month)

	date_day = date_format.day
	if date_day < 10:
		date_day = '0' + str(date_day)
	else:
		date_day = str(date_day)

	date_year = str(date_format.year)
	date_out = date_year + '_' + date_month + '_' + date_day

	return date_out

# function to rename CSV file
def name_data(data, data_type):
	date_col = data.apply(add_date, axis=1)

	data.insert(0, 'date_col', date_col)
	data = data.sort_values('date_col')

	start_date = data.head(1)['date_col'].values[0]
	end_date = data.tail(1)['date_col'].values[0]

	output_name = data_type + '_' + str(start_date) + '-' + str(end_date) + '.csv'

	return output_name

# load database file and find latest item date
def date_list():
	print('determining dates...')
	# find beginning dates
	purchase_data = pd.read_excel(database_file, sheet_name = 'Purchase Data')
	savings_data = pd.read_excel(database_file, sheet_name = 'Savings Data')

	purchase_date_string = purchase_data.tail(1)['Date'].iloc[0] # latest date
	savings_date_string = savings_data.tail(1)['Date'].iloc[0] # latest date

	# determine earliest date
	purchase_date = datetime.datetime.strptime(purchase_date_string, '%m/%d/%Y')
	savings_date = datetime.datetime.strptime(savings_date_string, '%m/%d/%Y')

	if savings_date <= purchase_date:
		latest_date = savings_date
	else:
		latest_date = purchase_date

	# find todays date
	end_date = datetime.datetime.today()
	end_date_string = end_date.strftime('%m/%d/%Y')

	# loop for date intervals
	start_date = latest_date - timedelta(days=30)
	start_date_string = start_date.strftime('%m/%d/%Y')

	mid_date = start_date + timedelta(days=round(rand_time(50,70)))

	date_list = []
	date_list.append([start_date.strftime('%m/%d/%Y'),mid_date.strftime('%m/%d/%Y')])
	while mid_date < end_date:
		start_date = mid_date
		mid_date = start_date + timedelta(days=round(rand_time(50,70)))
		start_date = start_date - timedelta(days=round(rand_time(20,30)))
		if mid_date > end_date:
			date_list.append([start_date.strftime('%m/%d/%Y'),end_date.strftime('%m/%d/%Y')])
		else:
			date_list.append([start_date.strftime('%m/%d/%Y'),mid_date.strftime('%m/%d/%Y')])
	
	return date_list

# get all csvs for one account
def loop_csv(data_type):
	for dates in date_list:
		start_date_string = dates[0]
		end_date_string = dates[1]
		
		download_button = ''
		time.sleep(rand_time(0.5,2.0))
		while download_button == '':
			try:
				download_button = browser.find_element_by_link_text('DOWNLOAD')
			except:
				download_button = ''

		time.sleep(rand_time(0.5,2.0))
		download_button.click()

		time.sleep(rand_time(0.5,2.0))
		dropdown_button = ''
		while dropdown_button == '':
			try:
				time.sleep(rand_time(0.5,2.0))
				dropdown_button = browser.find_elements_by_xpath("//span[@class='k-dropdown-wrap k-state-default']")
				assert len(dropdown_button) >= 2
				assert dropdown_button[len(dropdown_button) - 1].text == '-- Select --'
				dropdown_button = dropdown_button[len(dropdown_button) - 1]
			except:
				dropdown_button = ''

		time.sleep(rand_time(0.5,2.0))
		dropdown_button.click()

		dropdown_csv = ''
		while dropdown_csv == '':
			try:
				dropdown_csv = browser.find_elements_by_xpath("//*[contains(text(),'CSV')]")
			except:
				dropdown_csv = ''

		dropdown_csv_button = ''
		for i in range(0, len(dropdown_csv)):
			try:
				assert dropdown_csv[i].text == 'CSV - Comma Separated Value'
				dropdown_csv_button = dropdown_csv[i]
			except:
				pass

		if dropdown_csv_button == '':
			pdb.set_trace()

		time.sleep(rand_time(0.5,2.0))
		dropdown_csv_button.click()

		download_button = browser.find_element_by_xpath("//button[@class='button orange account-downloadtransactions-download-button']")
		start_time = browser.find_element_by_xpath("//input[@data-val-greaterthan-field='MinDateRangeFrom']")
		end_time = browser.find_elements_by_xpath("//input[@data-val-lessthanorequal='Ending date must be earlier than today.']")[1]

		start_time.clear()
		end_time.clear()

		start_time.send_keys(start_date_string)
		end_time.send_keys(end_date_string)

		time.sleep(rand_time(0.5,2.0))

		download_button.click()

		#loop until download finished		
		while not os.path.exists(src):
			time.sleep(0.25)

		time.sleep(0.5)

		# generate file name
		try:
			data = pd.read_csv(src)
		except:
			print('error.')
			pdb.set_trace()

		output_name = name_data(data, data_type)
		dst = download_folder + '\\' + output_name

		#move transactions file
		try:
			os.rename(src, dst)
		except:
			print('error.')
			pdb.set_trace()

		time.sleep(0.5)

# determine date_list
date_list = date_list()

# check if there really is a need to update
if len(date_list) < 2:
	pdb.set_trace()

# get webpage
print('logging in...')
web_page = site_name
profile = profile_loc

browser = webdriver.Firefox(profile)
browser.get(web_page)

loginElem = browser.find_element_by_name('user')
loginElem.send_keys(user_name)

time.sleep(rand_time(0.5,2.0))
passwordElem = browser.find_element_by_name('password')
passwordElem.send_keys(user_pass)
time.sleep(rand_time(0.5,2.0))

loginButton = browser.find_element_by_css_selector('input[title=\'Sign Into Online Banking\']')
loginButton.click()

# get information
time.sleep(rand_time(1.0,2.0))

# https://stackoverflow.com/questions/34759787/fetch-all-href-link-using-selenium-in-python
a_list = []
update_num = 0
while len(a_list) == 0:
	a_list = browser.find_elements_by_xpath("//a[@href]")
	update_num = update_num + 1
	update_string = '\r' + str(update_num) + ' ' + str(len(a_list))
	sys.stdout.write(update_string)
	sys.stdout.flush()

print('')
a_list = browser.find_elements_by_xpath("//a[@href]")
checking_account = ''
savings_account = ''
credit_account = ''

while credit_account == '':
	try:
		credit_account = browser.find_element_by_link_text(credit_name)
	except:
		credit_account = ''

time.sleep(rand_time(1.0,2.0))

accountButton = browser.find_element_by_css_selector('a[id=\'lnkAccountSummary\']')

# credit account
data_type = 'credit'
print('getting credit...')
time.sleep(rand_time(1.0,2.0))
credit_account.click()
loop_csv(data_type)

# checking account
data_type = 'checking'
print('getting checking...')
time.sleep(rand_time(0.5,2.0))
accountButton = browser.find_element_by_css_selector('a[id=\'lnkAccountSummary\']')
accountButton.click()
time.sleep(rand_time(0.5,2.0))
checking_account = browser.find_element_by_link_text(checking_name)
checking_account.click()
loop_csv(data_type)

# savings account
data_type = 'savings'
print('getting savings...')
time.sleep(rand_time(0.5,2.0))
accountButton = browser.find_element_by_css_selector('a[id=\'lnkAccountSummary\']')
accountButton.click()
time.sleep(rand_time(0.5,2.0))
savings_account = browser.find_element_by_link_text(savings_name)
savings_account.click()
loop_csv(data_type)

# sign output_name
time.sleep(rand_time(0.5,2.0))
signOutButton = browser.find_element_by_css_selector('a[id=\'SingOutLnk\']')
signOutButton.click()
time.sleep(rand_time(0.5,2.0))

# close browser
browser.close()

print('done.')