# bank-statement-analysis
These python files can be used to better understand how you're spending your money.

statement_analysis.py will gather your bank statement CSV files, and combine them into a single excel file with a spending graph, spending pie, and multiple transaction lists.

file description:

bank_database.py - This python script will generate an excel analysis of the statement data in your bank account.

bank_scrape.py - This python script will download and rename bank statement CSVs from a bank account.

lookup_table.xlsx - This excel file is a lookup table to classify your transactions into categories. You need to customize this file with your own transaction types.

modules used: Pandas, NumPy, XlsxWriter, openpyxl, datetime, selenium
