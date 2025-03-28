#%% #INFO KNOWN BUGS
    # todo If split is done on most recent tranction, program will not identify last transaction

#%% #INFO LIBRARY IMPORTS AND SETTING

import xlwings as xw
import warnings
import sys
from pathlib import Path

import pandas as pd # version 2.0.3
import numpy as np 
import os
from datetime import datetime
import polars as pl
import re

warnings.simplefilter('ignore')
#%% #INFO SETTING ACTIVE WORKBOOK, WORKBOOK POINTERS, AND WORKBOOK VARIABLES
shtTrans = xw.sheets('ChaseTransactions')
zPending = 'z-Pending'
ExpenseTbl = shtTrans.tables["Table22"].range
NcomeTbl = shtTrans.tables["Table26"].range
LastBalance = float(shtTrans.range('CBalance').value)
LastProcessedFile = shtTrans.range('LastReadFilename').value

# Find all Chase bank CSV files matching the pattern
file_pattern = r'Chase2517_Activity_\d{8}\.CSV'
print(f"{list(Path('.').iterdir())=}")
files =  [f for f in Path('.').iterdir() if re.match(file_pattern, f.name)]
print(f"{files=}\n")
if not files:
    sys.exit("No matching Chase bank CSV files found.")

# Get the file with the most recent date in its name
ChaseFile = str(max(files, key=lambda x: datetime.strptime(x.stem.split('_')[-1], '%Y%m%d')))

# Check if file was already processed
if ChaseFile == LastProcessedFile:
    sys.exit("Most recent Chase file has already been processed.")

#%% #INFO SET INCOME,EXPENSE, PENDING TRANSACTION DFs AND LAST POSITNG DATE
#? polars dataframe is not an option for as a converter, 
#? a custom converter can be made, however, it's much easier to use polars built-in converter
ETransactions = pl.from_pandas(shtTrans.range(ExpenseTbl.address).options(pd.DataFrame, header=1, index=False).value)
NTransactions = pl.from_pandas(shtTrans.range(NcomeTbl.address).options(pd.DataFrame, header=1, index=False).value)

# step iso pending transactions
E_Pending=ETransactions.filter(pl.col('Description').str.contains(zPending)).select(pl.exclude('Posting Date'))
N_Pending=NTransactions.filter(pl.col('Description').str.contains(zPending)).select(pl.exclude('Posting Date'))

# step clean description column
for i, df in enumerate([E_Pending, N_Pending]):
    modified_df = df.with_columns(
        pl.col('Description').str.strip_chars() #? remove leading and trailing spaces
            .str.slice(20) #? remove first 20 characters
            .str.replace(r'.{4}$', '') #? remove last 4 characters
            .str.replace_all('\s+',' ') #? remove multiple spaces
            .str.replace_all('\*','') #? remove asterisks
    )
    if i == 0:
        E_Pending_dict = modified_df.to_pandas().to_dict('index')
    else:
        N_Pending_dict = modified_df.to_pandas().to_dict('index')

# step remove pending transactions
ETransactions = ETransactions.filter(~pl.col('Description').str.contains(zPending))
NTransactions = NTransactions.filter(~pl.col('Description').str.contains(zPending))

#  step retrieve last posting date
LastPostDate = max(ETransactions['Posting Date'].max(), NTransactions['Posting Date'].max())



#%% #INFO READ NEW FILE AND SET IMPORTANT DATA TYPES
file = pl.scan_csv(ChaseFile, truncate_ragged_lines=True).with_columns( #? clean Blance and Posting Date columns
        pl.col('Balance').cast(float, strict=False),
        pl.col('Posting Date').str.strptime(pl.Datetime, '%m/%d/%Y')
    )
#%% #INFO GET AND TRANSFORM NEW TRANSACTIONS AND UPDATE NEW BALANCE
NewTrans = file.with_row_count().filter( #? filter out old transactions
        pl.col('row_nr') < (
            file.with_row_count().filter(
                (pl.col('Balance') == LastBalance) &
                (pl.col('Posting Date') == LastPostDate)
            ).collect().item(row=0, column='row_nr')
        )
    ).with_columns( #? add new columns amd modify Description column
        (pl.when(pl.col('Balance').is_null()
            ).then(pl.lit(zPending) + pl.col('Description')
            ).otherwise(pl.col('Description')
            ).alias('Description'))
                .str.replace_all('\s+',' ')
                .str.replace_all('\*',''),
        pl.lit(None).alias('Category'),
        pl.lit(None).alias('Sub-Category'),
        pl.col('Posting Date').dt.truncate("1mo").alias('Pay Period'),  # Always truncate to 1st of month
    ).collect()

if NewTrans.is_empty():
    sys.exit()

# Get the first non-null balance if it exists, otherwise use LastBalance
NewPostedTrans = NewTrans.filter(pl.col('Balance').is_not_null())['Balance']
NewBalance = NewPostedTrans[0] if not NewPostedTrans.is_empty() else LastBalance
    
#%% #INFO SET AUTOTAG DICTIONARY
shtTag = xw.sheets('Autotag')
Autotag = shtTag.range('a1').expand().options(pd.DataFrame, header=1, index=1).value
Autotag_dict = Autotag.to_dict('index')

#%% #INFO AUTOTAG NEW TRANSACTIONS FROM OLD PENIDNG TRANSACTIONS
for dct in [E_Pending_dict , N_Pending_dict]:
    for catch, values in dct.items():
        
        Temp = NewTrans.filter(
            (pl.col('Description').str.to_lowercase().str.contains(values['Description'].lower())) &
            (pl.col('Category').is_null()) &
            (pl.col('Amount') == values['Amount'])
        ).select(pl.col(['Category', 'Sub-Category', 'Pay Period', 'row_nr']))
        
        if not Temp.is_empty():
            for col in ['Category', 'Sub-Category', 'Pay Period']:
                try:
                    NewTrans = NewTrans.with_columns(
                        (pl.when(pl.col('row_nr').is_in(Temp.get_column('row_nr'))
                            ).then(values[col])
                            ).otherwise(pl.col(col)
                            ).alias(col)
                    )
                except:
                    pass
#%% #INFO AUTOTAG NEW TRANSACTIONS FROM AUTOTAG SHEET
for col in ['Category', 'Sub-Category']:
    for catch in Autotag_dict.keys():
        # print(f'{col=}\n{catch=}\n{Autotag_dict[catch][col]=}\n{pl.col(col)=}')
        NewTrans = NewTrans.with_columns(
            (pl.when(pl.col('Description').str.to_lowercase().str.contains(catch.lower()) 
                     & pl.col(col).is_null()
                ).then(pl.lit(Autotag_dict[catch][col]))
                ).otherwise(pl.col(col)
                ).alias(col)
        )

NewTrans = NewTrans.with_columns(
        pl.when(pl.col('Category').is_not_null())
            .then(pl.col('Description') + ' (A)')
            .otherwise(pl.col('Description'))
            .alias('Description')
    )

NewTransExpse = NewTrans.filter(pl.col('Amount') <= 0).select(['Pay Period', 'Posting Date', 'Description', 'Amount', 'Category', 'Sub-Category'])
NewTransNcome = NewTrans.filter(pl.col('Amount') > 0).select(['Pay Period', 'Posting Date', 'Description', 'Amount', 'Category'])

# %% #INFO UPDATE TRANSACTION SHEET WITH NEW TRANSACTIONS AND BALANCE

for tbl, dList, New in zip([ExpenseTbl, NcomeTbl], [len(E_Pending_dict), len(N_Pending_dict)], [NewTransExpse, NewTransNcome]):
    if dList != 0:
        shtTrans.range(tbl(2,1).address,tbl(dList+1,len(New.columns)).address).delete(shift='up') #? delete pending transactions

    if len(New) != 0:
        NewRecords = shtTrans.range(tbl(2,1).address,tbl(len(New)+1,len(New.columns)).address) #?define range for new records
        NewRecords.insert(shift = 'down', copy_origin='format_from_right_or_below') #? insert spacce for new records

        shtTrans.range(tbl(len(New)+2,1).address, tbl(len(New)+2,len(New.columns)).address).color = (169, 208, 142) #? marker for new record insert
        NewRecords = shtTrans.range(tbl(1,1).address,tbl(len(New)+1,len(New.columns)).address) #? redefine range for new records
        NewRecords.color = None 
        NewRecords.options(pd.DataFrame,header=1, index=False).value = New.to_pandas()


shtTrans.range('CBalance').value = NewBalance

shtTrans.range('LastReadFilename').value = ChaseFile

#%% #INFO CLEANUP PROCESSED FILES
# Delete the current file and any remaining Chase files
for f in files:  # Reuse the already matched files
    try:
        Path(f).unlink()
    except:
        pass

#%%


