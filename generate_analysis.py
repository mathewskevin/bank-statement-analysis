# Kevin Mathews 5/03/2019 rev 1.02
# Bank Statement Analysis Script
# written in Python 3

# Account statement data can be downloaded multiple chunks as CSV files.
# The following code will generate an excel file.
# This file will combine multiple bank statement CSV files into one list of transactions.
# The file will also include line/pie graphs to analyze your bank statement.
# The original CSV files will remain unmodified. Keep them for your records.

print('starting...')

import os, pdb, datetime
import numpy as np
import pandas as pd
import pdb
import xlsxwriter
from openpyxl import Workbook

# specify folder where CSV files are stored
data_folder = 'dummy_data'

# specify lookup table for transaction classification
# please update your own lookup table with transactions/categories
lookup_file = 'lookup_table.xlsx'

# specify output file name
output_name = 'generated_file (please enable editing).xlsx'

data_dict = {'credit':'Transaction Date',
			 'checking':'Date',
			 'savings':'Date'}

# function to convert date to quarter (not used)
def quarter_col(row):
		
	quarter_dict = {'01':'01',
				  '02':'01',
				  '03':'01',
				  '04':'02',
				  '05':'02',
				  '06':'02',
				  '07':'03',
				  '08':'03',
				  '09':'03',
				  '10':'04',
				  '11':'04',
				  '12':'04',}

	month_num = row['MM']
	year_num = str(row['Year'])[~1:]
	quarter_name = quarter_dict[month_num]
	
	output_string = '20' + year_num + quarter_name

	return output_string

# function to generate month column (not used)
def month_col(row):
	month_dict = {'01':'Jan',
				  '02':'Feb',
				  '03':'Mar',
				  '04':'Apr',
				  '05':'May',
				  '06':'Jun',
				  '07':'Jul',
				  '08':'Aug',
				  '09':'Sep',
				  '10':'Oct',
				  '11':'Nov',
				  '12':'Dec',}

	month_num = row['MM']
	year_num = str(row['Year'])[~1:]
	month_name = month_dict[month_num]

	output_string = '20' + year_num + month_num
	
	return output_string

# function to add date column for sorting
def add_date(row):
	data_type = row['Account Type']
	date_val = row[data_dict[data_type]]
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

# function to clean checking account data
def checking_convert(data):
	debit_string = 'POS Debit - Visa Check Card 7355 - '

	try:
		combine_debit = data[data['Description'].str.contains(debit_string)]['Description'].str.split('- ',2,expand=True)[2]
	except:
		combine_debit = pd.DataFrame()
		
	combine_normal = data[~data['Description'].str.contains(debit_string)]['Description']
	new_desc = pd.concat([combine_debit, combine_normal]).sort_index()

	data_out = data.drop('Description', axis = 1)
	data_out.insert(loc=2, column='Description', value=new_desc)

	data_out = data_out.drop('No.', axis = 1)

	assert data_out.shape[0] == data.shape[0]
	
	return data_out

# function to clean credit account data
def credit_convert(data):
	data_out = data.drop('Posted Date', axis = 1) # drop posted date
	data_out.rename(columns={'Transaction Date':'Date'}, inplace=True)# rename transaction date column

	return data_out

# function to clean savings account data
def savings_convert(data):
	data_out = data.drop('No.', axis = 1)

	return data_out

print('cleaning data...')
file_list = os.listdir(data_folder)
file_df = pd.DataFrame({'Filenames':file_list})
lookup_table = pd.read_excel(lookup_file)

# https://www.geeksforgeeks.org/python-pandas-split-strings-into-two-list-columns-using-str-split/
new_series = file_df['Filenames'].str.split("_", n=1, expand = True)
file_df['Type'] = new_series[0]

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(output_name, engine='xlsxwriter')

final_dataset = pd.DataFrame()
for key, value in data_dict.items():
	file_df_sort = file_df[file_df['Type']==key]
	main_dataset = pd.DataFrame()
	for index, row in file_df_sort.iterrows():
		file_name = row['Filenames']
		data_read = pd.read_csv(data_folder + '\\' + file_name)

		main_dataset = pd.concat([main_dataset, data_read])

	main_dataset['Account Type'] = key
	date_col = main_dataset.apply(add_date, axis=1)
	main_dataset.insert(0, 'date_col', date_col)
	main_dataset = main_dataset.sort_values('date_col')
	main_dataset = main_dataset.drop_duplicates()
	
	main_dataset = main_dataset.reset_index(drop=True)
	
	# edit checking dataset
	if key == 'checking':
		main_dataset = checking_convert(main_dataset)
	
	# edit credit dataset
	if key == 'credit':
		main_dataset = credit_convert(main_dataset)
	
	# edit savings dataset
	if key == 'savings':
		main_dataset = savings_convert(main_dataset)
	
	final_dataset = pd.concat([final_dataset, main_dataset], sort=True)
	main_dataset = main_dataset.drop(['date_col', 'Account Type'],axis=1)

time_df = final_dataset['date_col'].str.split('_', expand=True).drop(2, axis=1)
time_df.columns = ['Year', 'MM']
time_df['Month'] = time_df.apply(month_col, axis=1)
time_df['Quarter'] = time_df.apply(quarter_col, axis=1)

final_dataset = pd.concat([time_df, final_dataset], axis=1)

final_dataset = final_dataset.merge(lookup_table, on='Description', how='outer')
final_dataset = final_dataset.sort_values('date_col')

try:
	final_dataset['Year'] = final_dataset['Year'].astype('int64')
except:
	pdb.set_trace()
final_dataset['Month'] = final_dataset['Month'].astype('int64')
final_dataset['Quarter'] = final_dataset['Quarter'].astype('int64')
final_dataset = final_dataset[['Year','Month','Date','Account Type','Category','Description','Debit','Credit']]

# purchase data
purchase_data = final_dataset[final_dataset['Account Type'].isin(['checking','credit'])]
purchase_data = purchase_data[~purchase_data['Debit'].isnull()]
purchase_data = purchase_data[purchase_data['Description']!='Transfer to Credit Card']
purchase_data = purchase_data.drop('Credit', axis=1)

# savings data
savings_data = final_dataset[final_dataset['Account Type']=='savings']

# chart data
chart_data = pd.pivot_table(savings_data, values=['Debit','Credit'], index='Month', aggfunc=np.sum).reset_index()
spending_data = pd.pivot_table(purchase_data, values=['Debit'], index='Month', aggfunc=np.sum).reset_index()

chart_data = pd.merge(chart_data, spending_data, how='inner', on='Month')
chart_data.columns = ['Month','Savings','Checking','Spent']

year_cutoff = 2018
pie_data = purchase_data[purchase_data['Year']>=year_cutoff]
#pie_data = purchase_data
pie_chart_data = pd.pivot_table(pie_data, values=['Debit'], index='Category', aggfunc=np.sum).reset_index().sort_values('Debit', ascending=False)
dashboard_data = pd.pivot_table(pie_data, values=['Debit'], index='Category', aggfunc='count').reset_index()
dashboard_data = pd.merge(dashboard_data, pie_chart_data, how='inner', on='Category')
dashboard_data.columns = ['Category','# Purchases','$ Amount']
dashboard_data = dashboard_data.sort_values('$ Amount', ascending=False)

# paste all data
dashboard_data.to_excel(writer, sheet_name='Dashboard', index=False, startrow=2, startcol=9)
purchase_data.to_excel(writer, sheet_name = 'Purchase Data', index=False)
savings_data.to_excel(writer, sheet_name = 'Savings Data', index=False)
pie_chart_data.to_excel(writer, sheet_name='Chart Data', index=False, startrow=0, startcol=8)
chart_data.to_excel(writer, sheet_name='Chart Data', index=False)

worksheet = writer.sheets['Chart Data']
for row in range(2, chart_data.shape[0] + 2): # sum of savings
	formula_string = '=SUM(' + '$B$2:$B' + str(row) + ')'
	worksheet.write_formula('E' + str(row), formula_string)

	formula_string = '=SUM(' + '$C$2:$C' + str(row) + ')' # sum of spending
	worksheet.write_formula('F' + str(row), formula_string)

	formula_string = '=E' + str(row) + '-F' + str(row)# account balance
	worksheet.write_formula('G' + str(row), formula_string)

workbook  = writer.book
worksheet = writer.sheets['Dashboard']
worksheet.write(1,9, 'Purchase Breakdown ' + str(year_cutoff) + ' - Now')
worksheet.set_column(9, 9, 18)
worksheet.set_column(10, 10, 10.11)
worksheet.set_column(11, 11, 10.11)

# Create a new chart object.
chart = workbook.add_chart({'type': 'line'})
paste_string_1 = '=\'Chart Data\'!$G$2:$G$' + str(chart_data.shape[0] + 2) # data
paste_string_2 = '=\'Chart Data\'!$A$2:$A$' + str(chart_data.shape[0] + 2) # labels
chart.add_series({'values': paste_string_1, 'categories': paste_string_2, 'name':'Savings'})

bar_chart = workbook.add_chart({'type': 'column'})
paste_string_1 = '=\'Chart Data\'!$D$2:$D$' + str(chart_data.shape[0] + 2) # labels
paste_string_2 = '=\'Chart Data\'!$A$2:$A$' + str(chart_data.shape[0] + 2) # data
bar_chart.add_series({'values': paste_string_1, 'categories': paste_string_2, 'name':'Spending'})

chart.set_title({'name': 'Bank Account'})
chart.set_x_axis({'name':'YYYYMM', 'label_position':'low'})
chart.set_y_axis({'name':'Dollars ($)'})

chart.combine(bar_chart)

# insert the chart into the worksheet.
worksheet.insert_chart('B2', chart)

# pie chart
pie_chart = workbook.add_chart({'type':'pie'})
val_string = '=\'Chart Data\'!$J$2:$J$' + str(pie_chart_data.shape[0] + 1)
cat_string = '=\'Chart Data\'!$I$2:$I$' + str(pie_chart_data.shape[0] + 1)
pie_chart.add_series({
    'name':       'Purchases ' + str(year_cutoff) + ' - Now',
    'categories': cat_string,
    'values':     val_string,
})

worksheet.insert_chart('B17', pie_chart)

sheet_names = ['Purchase Data','Savings Data']
for i in sheet_names:
	worksheet = writer.sheets[i]
	worksheet.set_column(0, 0, 4.22)
	worksheet.set_column(1, 1, 6.44)
	worksheet.set_column(2, 2, 9.78)
	worksheet.set_column(3, 3, 11.78)
	worksheet.set_column(4, 4, 18)
	worksheet.set_column(5, 5, 52.22)

worksheet = writer.sheets['Chart Data']
worksheet.set_column(8, 8, 18)

# save database file
writer.save()

print('done.')