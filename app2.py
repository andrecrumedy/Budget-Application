#%% #INFO LIBRARY IMPORTS AND SETTING
import xlwings as xw
import warnings
import sys
import pandas as pd
import numpy as np
from datetime import datetime
import polars as pl
import sys
from datetime import datetime
from pathlib import Path
import polars as pl
import re


warnings.simplefilter('ignore')

#%% #INFO SETTING ACTIVE WORKBOOK AND WORKBOOK POINTERS
shtTrans = xw.sheets('Transactions')
ExpenseTbl = shtTrans.tables["Table22"].range
NcomeTbl = shtTrans.tables["Table26"].range
LastProcessedFile = shtTrans.range('LastReadFilename').value

#%% #INFO SET INCOME,EXPENSE, AND LAST DATE TRANSACTION DFs
# Convert Excel tables to polars dataframes
ETransactions = pl.from_pandas(shtTrans.range(ExpenseTbl.address).options(pd.DataFrame, header=1, index=False).value)
NTransactions = pl.from_pandas(shtTrans.range(NcomeTbl.address).options(pd.DataFrame, header=1, index=False).value)

# Get last posting date from existing transactions
LastPostDate = max(ETransactions['Posting Date'].max(), NTransactions['Posting Date'].max())

#%% #INFO SET LAST DATE TRANSACTION DFs
# Extract transactions from LastPostDate
E_LastDate = ETransactions.filter(pl.col('Posting Date') >= LastPostDate)
N_LastDate = NTransactions.filter(pl.col('Posting Date') >= LastPostDate)

# Clean description column for LastDate transactions
for i, df in enumerate([E_LastDate, N_LastDate]):
    modified_df = df.with_columns(
        pl.col('Description').str.strip_chars()  # remove leading and trailing spaces
            .str.replace_all('\s+', ' ')  # remove multiple spaces
            .str.replace_all('\*', '')    # remove asterisks
    )
    if i == 0:
        E_LastDate_dict = modified_df.to_pandas().to_dict('index')
    else:
        N_LastDate_dict = modified_df.to_pandas().to_dict('index')

# Remove LastPostDate transactions from main dataframes
ETransactions = ETransactions.filter(pl.col('Posting Date') < LastPostDate)
NTransactions = NTransactions.filter(pl.col('Posting Date') < LastPostDate)

#%% #INFO READ NEW FILE AND SET IMPORTANT DATA TYPES
# Find all CSV files matching the pattern (date-date.csv)
file_pattern = r'\d{2}_\d{2}_\d{2}-\d{2}_\d{2}_\d{2}\.csv'
files =  [f for f in Path('.').iterdir() if re.match(file_pattern, f.name)]

if not files:
    sys.exit("No matching CSV files found.")

# Get the file with the most recent date in its name
latest_file = max(files, key=lambda x: datetime.strptime(x.stem.split('-')[0], '%d_%m_%y'))

# Check if file was already processed
if str(latest_file) == LastProcessedFile:
    sys.exit("Most recent transaction file has already been processed.")

# Read the latest file
file = pl.scan_csv(latest_file, truncate_ragged_lines=True).with_columns(
    pl.col('Amount').str.replace('$', '').str.replace(',', '').cast(float),
    pl.col('Transaction Date').str.strptime(pl.Datetime, '%m/%d/%Y')
)

#%% #INFO GET AND TRANSFORM NEW TRANSACTIONS
NewTrans = file.filter(
    pl.col('Transaction Date') >= LastPostDate
).with_columns(
    pl.coalesce(pl.col('Description'), pl.col('Type'))
        .str.strip_chars()
        .str.replace_all('\s+', ' ')
        .str.replace_all('\*', '')
        .alias('Description'),
    pl.lit(None).alias('Category'),
    pl.lit(None).alias('Sub-Category'),
    pl.col('Transaction Date').dt.truncate("1mo").alias('Pay Period'),
    pl.col('Transaction Date').alias('Posting Date')
).select([
    'Pay Period',
    'Posting Date',
    'Description',
    'Amount',
    'Category',
    'Sub-Category'
]).collect()

if NewTrans.is_empty():
    sys.exit()

#%% #INFO SET AUTOTAG DICTIONARY
shtTag = xw.sheets('Autotag')
Autotag = shtTag.range('a1').expand().options(pd.DataFrame, header=1, index=1).value
Autotag_dict = Autotag.to_dict('index')

#%% #INFO AUTOTAG NEW TRANSACTIONS FROM LAST DATE TRANSACTIONS
for dct in [E_LastDate_dict, N_LastDate_dict]:
    for catch, values in dct.items():
        Temp = NewTrans.filter(
            (pl.col('Description').str.to_lowercase().str.contains(values['Description'].lower())) &
            (pl.col('Category').is_null()) &
            (pl.col('Amount') == values['Amount']) &
            (pl.col('Posting Date') == values['Posting Date'])
        )
        
        if not Temp.is_empty():
            for col in ['Category', 'Sub-Category', 'Pay Period']:
                try:
                    NewTrans = NewTrans.with_columns(
                        pl.when(pl.col('Description').str.to_lowercase().str.contains(values['Description'].lower()) &
                               (pl.col('Amount') == values['Amount']) &
                               (pl.col('Posting Date') == values['Posting Date']))
                        .then(values[col])
                        .otherwise(pl.col(col))
                        .alias(col)
                    )
                except:
                    pass

#%% #INFO AUTOTAG NEW TRANSACTIONS FROM AUTOTAG SHEET
for col in ['Category', 'Sub-Category']:
    for catch in Autotag_dict.keys():
        NewTrans = NewTrans.with_columns(
            pl.when(pl.col('Description').str.to_lowercase().str.contains(catch.lower()) 
                    & pl.col(col).is_null()
                ).then(pl.lit(Autotag_dict[catch][col]))
                .otherwise(pl.col(col))
                .alias(col)
        )

NewTrans = NewTrans.with_columns(
    pl.when(pl.col('Category').is_not_null())
        .then(pl.col('Description') + ' (A)')
        .otherwise(pl.col('Description'))
        .alias('Description')
)

# Split transactions into expense and income
NewTransExpse = NewTrans.filter(pl.col('Amount') <= 0).select(['Pay Period', 'Posting Date', 'Description', 'Amount', 'Category', 'Sub-Category'])
NewTransNcome = NewTrans.filter(pl.col('Amount') > 0).select(['Pay Period', 'Posting Date', 'Description', 'Amount', 'Category'])

#%% #INFO UPDATE TRANSACTION SHEET WITH NEW TRANSACTIONS
for tbl, dList, New in zip([ExpenseTbl, NcomeTbl], 
                          [len(E_LastDate_dict), len(N_LastDate_dict)], 
                          [NewTransExpse, NewTransNcome]):
    if dList != 0:
        # Delete transactions from LastPostDate
        shtTrans.range(tbl(2,1).address,tbl(dList+1,len(New.columns)).address).delete(shift='up')

    if len(New) != 0:
        NewRecords = shtTrans.range(tbl(2,1).address,tbl(len(New)+1,len(New.columns)).address)
        NewRecords.insert(shift = 'down', copy_origin='format_from_right_or_below')

        shtTrans.range(tbl(len(New)+2,1).address, tbl(len(New)+2,len(New.columns)).address).color = (169, 208, 142)
        NewRecords = shtTrans.range(tbl(1,1).address,tbl(len(New)+1,len(New.columns)).address)
        NewRecords.color = None 
        NewRecords.options(pd.DataFrame,header=1, index=False).value = New.to_pandas()

shtTrans.range('LastReadFilename').value = str(latest_file)

#%% #INFO CLEANUP PROCESSED FILES
# Delete the current file and any remaining transaction files
# for f in files: 
#     try:
#         Path(f).unlink()
#     except:
#         pass