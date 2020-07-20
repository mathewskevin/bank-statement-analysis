# bank-statement-analysis
I use these python files to aggregate my bank statements (CSV files), and combine them into a single excel file with a spending graph, spending pie, and multiple transaction lists. Makes it easy to budget.

file description:

bank_scrape.py - A selenium webscraper which I use to scrape my bank's website for all bank statements for all cards/accounts. It'll check what statements I already have and download the newest data. It's customized to my bank's web interface. No issues with race conditions or security so far.

bank_database.py - This python script generates the excel analysis given the statement data in a folder on my computer.

lookup_table.xlsx - This file contains some extra data used by bank_database.py. I can add custom notes to certain transactions. This way I remember exactly why I spent $400 on 27 copies of Deep Impact on DVD.
