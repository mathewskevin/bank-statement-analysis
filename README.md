# navyfed-account-analysis
These python files can be used to better understand how you're spending your money.

navyfed_database.py will gather your NavyFed CSV files (downloaded from navyfederal.org), and combine them into a single excel file with a spending graph, spending pie, and multiple transaction lists.

file description:

navyfed_database.py - This python script will generate an excel analysis of the statement data in your NavyFed account.

navyfed_name_csv.py - This python script will rename NavyFed CSVs with unique filenames

lookup_table.xlsx - This excel file is a lookup table to classify your transactions into categories. You need to customize this file with your own transaction types.

modules used: Pandas, NumPy, XlsxWriter, openpyxl, datetime

https://www.navyfederal.org/
