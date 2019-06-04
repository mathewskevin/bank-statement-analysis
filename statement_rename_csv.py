# Kevin Mathews 6/3/2019
# Bank Statement CSV Naming Script
# written in Python 3

# This script will rename a bank statement CSV file (e.g. transactions.csv), with a unique name.
# Please collect your CSV files into one folder. 

import numpy as np
import pandas as pd
import os, datetime
import pdb

# specify type of account (e.g. savings, checking, credit, etc.)
# this will control how the script interprets the file
data_type = 'savings'

# specify name of input CSV, place in same directory as python file
input_filename = 'transactions.csv'

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
def name_data(data):
	date_col = data.apply(add_date, axis=1)

	data.insert(0, 'date_col', date_col)
	data = data.sort_values('date_col')

	start_date = data.head(1)['date_col'].values[0]
	end_date = data.tail(1)['date_col'].values[0]

	output_name = data_type + '_' + str(start_date) + '-' + str(end_date) + '.csv'

	return output_name
	
data = pd.read_csv(input_filename)
output_name = name_data(data)

# write new CSV file
os.rename(input_filename, output_name)