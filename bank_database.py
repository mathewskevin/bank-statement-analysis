# Kevin Mathews 09/14/2020 rev 1.04
# Bank Analysis Script
# written in Python 3

# Account statement data can be downloaded from Bank Site in 90 day chunks as CSV files.
# The following code will generate an excel file.
# This file will combine multiple bank statement CSV files into one list of transactions.
# The file will also include line/pie graphs to analyze your bank statement.
# The original CSV files will remain unmodified. Keep them for your records.

# X Object Oriented Structure
# X Inflow & Outflow DFs
# X Labels
# Purchase Table
# Graphs over time

import os, sys, datetime
import numpy as np
import pandas as pd
import pdb
import xlsxwriter
from tabulate import tabulate
from openpyxl import Workbook
import config_database

# year cutoff
year_cutoff = 2010

# specify folder where CSV files are stored
data_folder = os.path.join(os.getcwd(), 'bank_files')

# specify lookup table for transaction classification
# please update your own lookup table with transactions/categories
#lookup_filename = os.path.join(data_folder, 'database_lookup.xlsx')
lookup_filename = os.path.join(os.getcwd(), 'database_lookup.xlsx')

# specify output file name
output_name = 'database_analysis.xlsx'

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
                    '12':'04'}

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
                  '12':'Dec'}

    month_num = row['MM']
    year_num = str(row['Year'])[~1:]
    month_name = month_dict[month_num]

    output_string = '20' + year_num + month_num
    
    return output_string

# https://stackoverflow.com/questions/23861680/convert-spreadsheet-number-to-column-letter
def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def gen_month_list(start_month, end_month):
    assert start_month < end_month

    # find current month
    current_month_str = str(start_month)[~1:]
    current_year_str = str(start_month)[:4]

    month_output_list = [start_month]
    month_list = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    current_month = start_month
    while current_month != end_month:
        
        # find next month
        current_month_idx = month_list.index(current_month_str)

        if current_month_idx == 11:
            next_month_idx = 0
            current_year_str = str(int(current_year_str) + 1)
        
        else:
            next_month_idx = current_month_idx + 1

        current_month_str = month_list[next_month_idx]
        current_month = int(current_year_str + current_month_str)

        # add next month to list
        month_output_list.append(current_month)

    #print(month_output_list)
    return month_output_list

class account_template():
    # function to add date column for sorting
    def add_time_cols(self):
        def add_date(row):
            # date conversion dictionary
            data_dict = {'Credit Account':'Transaction Date',
                         'Checking Account':'Date',
                         'Savings Account':'Date'}

            date_val = row[data_dict[self.account_type]]
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

        mid_dataset = self.df.copy()

        date_col = mid_dataset.apply(add_date, axis=1)
        mid_dataset.insert(0, 'date_col', date_col)
        mid_dataset = mid_dataset.sort_values('date_col', ascending=False)

        # remove duplicates & fix index
        mid_dataset = mid_dataset.drop_duplicates()
        mid_dataset = mid_dataset.reset_index(drop=True)

        # time dataframe
        time_df = mid_dataset['date_col'].str.split('_', expand=True).drop(2, axis=1)
        
        time_df.columns = ['Year', 'MM']
        time_df['Month'] = time_df.apply(month_col, axis=1)
        #time_df['Quarter'] = time_df.apply(quarter_col, axis=1)

        # main dataset (holds earnings, spending, etc.)
        mid_dataset = pd.concat([time_df, mid_dataset], axis=1)
                
        mid_dataset['Year'] = mid_dataset['Year'].astype('int64')
        mid_dataset['Month'] = mid_dataset['Month'].astype('int64')
        #mid_dataset['Quarter'] = mid_dataset['Quarter'].astype('int64')

        self.df = mid_dataset.copy()

    # find inflow & outflow
    def calc_flow(self):
        data = self.df.copy()

        data.rename(columns={'Credit':'Inflow', 'Debit':'Outflow'}, inplace=True)
        #data.drop('date_col', axis=1, inplace = True)
        #data.drop('Year', axis=1, inplace = True)
        data.drop('MM', axis=1, inplace = True)

        self.df = data.copy()

        df_inflow = self.df[~self.df['Inflow'].isna()].copy()
        df_inflow.drop('Outflow', axis=1, inplace = True)
        df_outflow = self.df[~self.df['Outflow'].isna()].copy()
        df_outflow.drop('Inflow', axis=1, inplace = True)

        self.df_inflow = df_inflow.copy()
        self.df_outflow = df_outflow.copy()
        self.inflow = round(self.df_inflow['Inflow'].sum(), 2)
        self.outflow = round(self.df_outflow['Outflow'].sum(), 2)
        self.balance = round(self.inflow - self.outflow, 2)

    # add account name to inflow/outflow df
    def add_account_name(self):

        account_series = pd.Series([self.account_name] * self.df_inflow.shape[0])
        self.df_inflow.reset_index(inplace=True, drop=True)
        self.df_inflow.insert(0, 'Account', account_series)

        account_series = pd.Series([self.account_name] * self.df_outflow.shape[0])
        self.df_outflow.reset_index(inplace=True, drop=True)
        self.df_outflow.insert(0, 'Account', account_series)

class account_savings(account_template):

    # function to clean credit account data
    def df_update(self):
        data = self.df.copy()
        data_out = data.drop('No.', axis = 1)
        self.df = data_out.copy()

    def __init__(self, df, account_name):
        self.df = df; self.account_name = account_name
        self.account_type = 'Savings Account'
        self.add_time_cols()
        self.df_update()
        self.calc_flow()
        self.add_account_name()

class account_credit(account_template):

    # function to clean credit account data
    def df_update(self):
        data = self.df.copy()
        data_out = data.drop('Posted Date', axis = 1) # drop posted date
        data_out.rename(columns={'Transaction Date':'Date'}, inplace=True) # rename transaction date column
        self.df = data_out.copy()

    def __init__(self, df, account_name):
        self.df = df; self.account_name = account_name
        self.account_type = 'Credit Account'
        self.add_time_cols()
        self.df_update()
        self.calc_flow()
        self.add_account_name()
        
class account_checking(account_template):

    # function to clean checking account data
    def df_update(self):
        data = self.df.copy()
        data_cols = data.columns.values
        debit_string = config_database.debit_string

        combine_debit = data[data['Description'].str.contains(debit_string)]['Description'].str.split('- ',2,expand=True)[2]

        #try:
        #    combine_debit = data[data['Description'].str.contains(debit_string)]['Description'].str.split('- ',2,expand=True)[2]
        #except:
        #    combine_debit = pd.DataFrame()
            
        combine_normal = data[~data['Description'].str.contains(debit_string)]['Description']
        new_desc = pd.concat([combine_debit, combine_normal]).sort_index()

        data_out = data.drop('Description', axis = 1)
        data_out.insert(loc=2, column='Description', value=new_desc)

        data_out = data_out[data_cols].drop('No.', axis = 1)

        assert data_out.shape[0] == data.shape[0]
        
        self.df = data_out.copy()

    def __init__(self, df, account_name):
        self.df = df; self.account_name = account_name
        self.account_type = 'Checking Account'
        self.add_time_cols()
        self.df_update()
        self.calc_flow()
        self.add_account_name()

def gen_transactions(credit_outflow, checking_outflow, savings_inflow):
    # DELIVERABLE Transactions List
    transactions = pd.concat([credit_outflow, checking_outflow, savings_inflow])
    categories_table = pd.read_excel(lookup_filename, sheet_name = 'Lookup') # get list of unique purchases
    unique_transactions = pd.read_excel(lookup_filename, sheet_name = 'Unique Transactions') # get list of unique purchases
    unique_transactions.rename(columns={'Category':'Category Override'}, inplace=True)
    transactions = transactions.merge(unique_transactions, on=['Account','Date','Description','Outflow','Inflow'], how='outer')
    transactions = transactions.merge(categories_table, on=['Description'], how='left')
    transactions.loc[~transactions['Category Override'].isna(),'Category'] = transactions['Category Override']
    transactions.sort_values('date_col', ascending=False, inplace=True)
    transactions.drop_duplicates(inplace=True)
    transactions['Category'] = transactions['Category'].fillna('Abnormal - New') # classify blanks as new
    transactions = transactions[['Account','Date','Description','Category','Inflow','Outflow','Memo','Month']]
    return transactions

def gen_transactions_time_series(transactions):
   # Create Time Series Data
    cur_date = datetime.datetime.now()
    cur_month = int(str(cur_date.year) + str(cur_date.month).zfill(2))
    month_list = gen_month_list(201601, cur_month)
    month_index = pd.DataFrame(month_list, columns = ['Month'])

    # Inflow
    pivot_inflow = pd.pivot_table(transactions[(~transactions['Inflow'].isna())], values=['Inflow'], index='Month', aggfunc=np.sum).reset_index()
    pivot_outflow = pd.pivot_table(transactions[(~transactions['Outflow'].isna())], values=['Outflow'], index='Month', aggfunc=np.sum).reset_index()

    # Four Walls Spending
    pivot_4walls = pd.pivot_table(transactions[(~transactions['Outflow'].isna()) & (transactions['Category'].str.contains('Four Walls'))], values=['Outflow'], index='Month', aggfunc=np.sum).reset_index()
    pivot_4walls.columns = ['Month','Four Walls']

    # Investing
    pivot_investing = pd.pivot_table(transactions[(~transactions['Outflow'].isna()) & (transactions['Category'].str.contains('Investing'))], values=['Outflow'], index='Month', aggfunc=np.sum).reset_index()
    pivot_investing.columns = ['Month', 'Investing']

    # Abnormal / Unknown / One-Time Purchases
    pivot_abnormal = pd.pivot_table(transactions[(~transactions['Outflow'].isna()) & (transactions['Category'].str.contains('Abnormal'))], values=['Outflow'], index='Month', aggfunc=np.sum).reset_index()
    pivot_abnormal.columns = ['Month', 'Abnormal']

    # Discretionary
    pivot_discretionary = pd.pivot_table(transactions[(~transactions['Outflow'].isna()) & (transactions['Category'].str.contains('Discretionary'))], values=['Outflow'], index='Month', aggfunc=np.sum).reset_index()
    pivot_discretionary.columns = ['Month', 'Discretionary']

    # DLEIVERABLE Transaction Time Series
    transactions_time_series = month_index.merge(pivot_inflow, on=['Month'], how='left')
    transactions_time_series = transactions_time_series.merge(pivot_outflow, on = ['Month'], how='left')
    transactions_time_series = transactions_time_series.merge(pivot_4walls, on = ['Month'], how='left')
    transactions_time_series = transactions_time_series.merge(pivot_discretionary, on = ['Month'], how='left')
    transactions_time_series = transactions_time_series.merge(pivot_abnormal, on = ['Month'], how='left')
    transactions_time_series = transactions_time_series.merge(pivot_investing, on = ['Month'], how='left')
    transactions_time_series.fillna(0, inplace=True)
    transactions_time_series['Saved (Raw)'] = transactions_time_series['Inflow'] - transactions_time_series['Outflow'] # Find Savings & Wealth
    transactions_time_series['Saved (Wealth)'] = transactions_time_series['Saved (Raw)'] + transactions_time_series['Investing']
    transactions_time_series['Saved (Raw)'] = transactions_time_series['Saved (Raw)'].cumsum()
    transactions_time_series['Saved (Wealth)'] = transactions_time_series['Saved (Wealth)'].cumsum()
    transactions_time_series['Spent (Wealth)'] = (transactions_time_series['Outflow'] - transactions_time_series['Investing']).cumsum()
    transactions_time_series['Spent (Month)'] = transactions_time_series['Four Walls'] + transactions_time_series['Discretionary'] + transactions_time_series['Abnormal']
    transactions_time_series['Saved (Month)'] = (transactions_time_series['Inflow'] - transactions_time_series['Spent (Month)'])
    transactions_time_series['Spent (6 Mo Rolling)'] = transactions_time_series['Spent (Month)'].rolling(6).mean()
    transactions_time_series.sort_values('Month', ascending=False, inplace=True)
    transactions_time_series.loc[transactions_time_series['Month']==cur_month, 'Spent (6 Mo Rolling)'] = 'N/A'

    return transactions_time_series

def gen_transactions_distributions(transactions):
    # DELIVERABLE Transaction Distribution

    inflow_pivot = pd.pivot_table(transactions[(~transactions['Inflow'].isna())], values=['Inflow'], index='Month', aggfunc=np.sum) # YYYYMM Data
    inflow_pivot = inflow_pivot.sort_index(ascending=False).T
    inflow_pivot = pd.DataFrame(inflow_pivot.values, columns=[i for i in inflow_pivot.columns])
    inflow_pivot.index = ['Inflow']

    transactions_pivot = pd.pivot_table(transactions[(~transactions['Outflow'].isna())], values=['Outflow'], index='Category', columns='Month', aggfunc=np.sum).reset_index() # YYYYMM Data
    transactions_distribution = pd.DataFrame()
    for idx, row in transactions_pivot.iterrows():
        row = row.reset_index(level=0, drop=True)
        index = row.index[1:].values
        values = row.values[1:]
        title = row.values[0]
        df = pd.DataFrame(values, index=index, columns=[title])
        transactions_distribution = pd.concat([transactions_distribution, df], axis=1)

    # sort months & transpose
    transactions_distribution = transactions_distribution.sort_index(ascending=False).T 

    # Find public & private transactions
    public_categories = pd.read_excel(lookup_filename, sheet_name = 'Public Categories') # get list of public category types
    public_categories = public_categories.loc[:,'Category'].to_list()
    tdist_public = transactions_distribution.loc[public_categories].copy()
    tdist_private = transactions_distribution.loc[~transactions_distribution.index.isin(public_categories)].copy()
    tdist_private = tdist_private.loc[~tdist_private.index.str.contains('Investing')]

    tdist_discretionary = pd.DataFrame([i for i in tdist_private.sum().values], index = [i for i in tdist_private.sum().index], columns = ['Discretionary | Other']).T
    tdist_public = pd.concat([tdist_public, tdist_discretionary, inflow_pivot])
    tdist_public.sort_index(inplace = True)

    #transactions_distribution_public = pd.concat([transactions_distribution_public, inflow_pivot])
    #transactions_distribution_private = pd.concat([transactions_distribution_private, inflow_pivot])
    #transactions_distribution.reset_index(inplace=True)
    #inflow_pivot.reset_index(inplace=True)

    return transactions_distribution, tdist_public

def add_distribution_comments(worksheet, dist_df, transactions):
    # https://stackoverflow.com/questions/23861680/convert-spreadsheet-number-to-column-letter
    def colnum_string(n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def short_string(string):
        len_string = 18
        if len(string) > len_string:
            return string[0:len_string] + '...'
        else:
            return string

    def gen_comment_table(col_name, row_name, transactions):
        comment_table_values = transactions.loc[(transactions['Category']==row_name) & (transactions['Month']==col_name)]
        col_type = 'Outflow'

        if row_name == 'Discretionary | Other':
            comment_table_values = transactions.loc[(transactions['Category'].str.contains('Discretionary')) & (transactions['Month']==col_name)]
            comment_table_values = comment_table_values[comment_table_values['Category']!='Discretionary | Eating Out']

        if row_name == 'Inflow':
            comment_table_values = transactions.loc[(transactions['Category'].str.contains('Inflow')) & (transactions['Month']==col_name)]
            col_type = 'Inflow'

        comment_table_counts = pd.DataFrame(comment_table_values['Description'].value_counts()).reset_index()
        comment_table_counts.columns = ['Description','Count']

        try:
            comment_table = comment_table_values.pivot_table(index='Description', values=col_type, aggfunc=np.sum).reset_index().sort_values(col_type, ascending=False)
        except:
            pdb.set_trace()
        comment_table = comment_table.merge(comment_table_counts, on='Description')
        comment_table['Description'] = comment_table['Description'].apply(short_string)

        percent_col = (comment_table[col_type]/comment_table[col_type].sum()).round(2)
        percent_col = percent_col * 100
        percent_col = pd.DataFrame(percent_col)
        percent_col.columns = ['Percent']
        comment_table.reset_index(drop=True); percent_col.reset_index(drop=True)
        comment_table = pd.concat([comment_table, percent_col], axis=1)

        avg_col = (comment_table[col_type]/comment_table['Count']).round(2)
        avg_col = pd.DataFrame(avg_col)
        avg_col.columns = ['Average']
        comment_table = pd.concat([comment_table, avg_col], axis=1)

        comment_table = comment_table[['Description',col_type,'Count','Average','Percent']]
        return comment_table

    # initial cleaning for processing
    dist_df.fillna('None', inplace=True)
    dist_df.replace(0, 'None', inplace=True)

    #dist_df.set_index('index', inplace=True)

    # get lists of cols & rows for pulling dataframe cell coordinates
    col_list = [i for i in dist_df.columns.values]
    row_list = [i for i in dist_df.index]

    # cycle thorugh all cells of dataframe & input comments where values exist
    print('adding comments...')
    for col_name, col_data in dist_df.iteritems():
        for row_name, cell_data in col_data.iteritems():
            row_num = row_list.index(row_name); col_num = col_list.index(col_name)
            cell_value = dist_df.loc[row_name, col_name]
            current_col_letter = colnum_string(col_num + 2)
            cell_name = current_col_letter + str(row_num + 2)
           
            # where cells are not empty, add comment
            if cell_value != 'None':
                sys.stdout.write(str(col_name) + ' ' + cell_name + '\r')
                sys.stdout.flush()
                comment_table = gen_comment_table(col_name, row_name, transactions)

                # % of total
                #assert round(cell_value, 2) == round(comment_table['Outflow'].sum(), 2) 
                #assert round(cell_value, 2) == round(comment_table['Inflow'].sum(), 2)

                # % of category
                
                comment_string = tabulate(comment_table, tablefmt='simple', headers='keys', showindex=False)
                comment_string = row_name + ' ' + str(col_name) + '\n' + comment_string
                worksheet.write_comment(cell_name, comment_string, {'x_scale': 3, 'y_scale': 3, 'font_name':'monospace', 'font_size': 8}) #  'font_size': 10 SimSun

    return worksheet

print('running...')
file_list = os.listdir(data_folder)
file_df = pd.DataFrame({'Filenames':file_list}) # convert to dataframe
file_df = file_df[file_df['Filenames'].str.contains('.csv')] # filter to csvs

# https://www.geeksforgeeks.org/python-pandas-split-strings-into-two-list-columns-using-str-split/
new_series = file_df['Filenames'].str.split("_", n=1, expand = True) # split names by "_"
file_df['Type'] = new_series[0] # add checking/saving/credit label

# get account types
data_list = file_df['Type'].drop_duplicates().to_list()

# cycle thorugh csvs and merge into main datasest
account_dict = {}
for account_type in data_list:
    file_df_sort = file_df[file_df['Type']==account_type]
    mid_dataset = pd.DataFrame()
    for index, row in file_df_sort.iterrows():
        file_name = row['Filenames']
        data_read = pd.read_csv(os.path.join(data_folder, file_name))
        mid_dataset = pd.concat([mid_dataset, data_read])

    # edit checking dataset
    if account_type == 'checking':
        account_mid = account_checking(mid_dataset, config_database.checking_name)

    # edit credit dataset
    if account_type == 'credit':
        account_mid = account_credit(mid_dataset, config_database.credit_name)

    # edit savings dataset
    if account_type == 'savings':
        account_mid = account_savings(mid_dataset, config_database.savings_name)

    account_dict[account_mid.account_name] = account_mid
    print(account_mid.balance, '|', account_mid.account_name)

# list used for debug
account_keys = [i for i in account_dict]

# generate transactions
savings_inflow = account_dict.get(config_database.savings_name).df_inflow
df_inflow = savings_inflow.copy()
credit_outflow = account_dict.get(config_database.credit_name).df_outflow
checking_outflow = account_dict.get(config_database.checking_name).df_outflow
checking_outflow = checking_outflow.loc[checking_outflow['Description']!='Transfer to Credit Card',:]

# Deliverables
transactions = gen_transactions(credit_outflow, checking_outflow, savings_inflow)
transactions_time_series = gen_transactions_time_series(transactions)
transactions_distribution, tdist_public = gen_transactions_distributions(transactions)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(output_name, engine='xlsxwriter')

tdist_public.reset_index(inplace=True)
tdist_public.to_excel(writer, sheet_name ='Distribution', index=False)
tdist_public.set_index('index', inplace=True)

transactions_time_series.to_excel(writer, sheet_name ='Time Series', index=False)
transactions['Category'] = transactions['Category'].replace('Abnormal - New','')
transactions.to_excel(writer, sheet_name ='Transactions', index=False)

worksheet = writer.sheets['Transactions']
worksheet.set_column(1, 1, 10)
worksheet.set_column(2, 2, 50)
worksheet.set_column(3, 3, 25)

# add comments to distribution
worksheet = writer.sheets['Distribution']
worksheet.set_column(0, 0, 25)
worksheet = add_distribution_comments(worksheet, tdist_public, transactions)

writer.save()
print('done.')