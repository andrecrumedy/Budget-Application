# %% #INFO KNOWN BUGS
# - If split is done on most recent tranction, program will not identify last transaction

# %% [markdown]
# Penidng Updates
# - Deleting chase file

# %% [markdown]
# ## <ins> Library Imports

# %%
display('<-------------------------Importing Libraries and Configuring Settings------------------------->')

import xlwings as xw
import warnings
import sys

# xw.Range("a1").value = "NO"

import pandas as pd # version 2.0.3
import numpy as np 
import os
from datetime import datetime

warnings.simplefilter('ignore')
# pd.reset_option("all")

#     pd.set_option('display.max_rows', None)
#     pd.set_option('display.max_columns', None)

pd.set_option('display.max_colwidth', None)

# %% [markdown]
# ## <ins> Computation Functions

# %%
def Compute_PayPeriod(PostDate):
    PostDate = pd.to_datetime(PostDate)
    if PostDate.day >= 15:
        return datetime(year=PostDate.year, month=PostDate.month, day= 15) 
    else:
        return datetime(year=PostDate.year, month=PostDate.month, day= 1)

# %%
date = Compute_PayPeriod()

# %% [markdown]
# ## <ins> Current Transactions Table

# %% [markdown]
# <font color='blue'> 
# **Importing Current Transactions Income and Expense Tables  
#     Drop Pending Records in Ncome and Expense Tables  
#     Last Balance Amount and Last Posting Date  
#     Displays** 
# <font color='black'>
# - Using the Active Sheet To Find Workbook

# %%
display('<-------------------------Importing and Scubbing Current Transactions------------------------->')
zPending = 'z-Pending'
shtTrans = xw.sheets('Transactions')

# %%
ExpenseTbl = shtTrans.tables["Table22"].range
NcomeTbl = shtTrans.tables["Table26"].range
ETransactions = shtTrans.range(ExpenseTbl.address).options(pd.DataFrame, header=1, index=False).value
NTransactions = shtTrans.range(NcomeTbl.address).options(pd.DataFrame, header=1, index=False).value
# display(NTransactions)
E_Pending = ETransactions[ETransactions['Description'].str.contains(zPending)]
N_Pending = NTransactions[NTransactions['Description'].str.contains(zPending)]

# %%
for df in [E_Pending, N_Pending]:
    df['Description'] = df['Description'].str[20:-4].str.strip().replace(["\s+"], ' ', regex = True)
    df.drop('Posting Date', axis='columns', inplace=True)
    df['Description'].replace({'\*':'',"\s+":' '}, regex=True, inplace=True)

# %%
# E_Pending['Description'].replace('\*','', regex=True)

# %%
E_Pending_dict = E_Pending.to_dict('index')
N_Pending_dict = N_Pending.to_dict('index')

# %%
# display(NTransactions)


# %%
N_Pending

# %%
eList = E_Pending.index.to_list()
nList = N_Pending.index.to_list()

for df, lst in zip([ETransactions, NTransactions], [eList, nList]):
    df.drop(lst, axis='index', inplace=True)
    df.reset_index(drop=True, inplace=True)

LastPostDate = max(ETransactions['Posting Date'].max(), NTransactions['Posting Date'].max())
LastBalance = float(shtTrans.range('CBalance').value)

display('Last Posting Date: ' + LastPostDate.strftime('%m/%d/%y'),'Last Balance: ' + str(LastBalance))
display('Showing Below Expense Table (without Pending)', ETransactions.head(), \
        'Showing Below Income Table (without Pending)', NTransactions.head())

# %% [markdown]
# ## <ins> New Transactions Table Part 1

# %% [markdown]
# <font color='blue'> 
#  **Import Updated Transactions  using file name in current folder and todays date  
#     Datatype transformations  
#     Identify Last Inserted Record (assuming it is in dataset)  
#     Identify New Balance  
#     Pre-fix new Transactions with 'z-Pending'  
#     Pull Needed Columns  
#     Remove White Space and Special Characters  
#     Displays**

# %%
display('<-------------------------Importing and Scubbing New Transactions------------------------->')
CurrentDateTime = datetime.now()
ChaseFile = 'Chase2517_Activity_' + CurrentDateTime.strftime('%Y%m%d') + '.CSV' 
NewTrans = pd.read_csv(ChaseFile, header = 0, index_col=False)

NewTrans['Balance'] = pd.to_numeric(NewTrans['Balance'], errors='coerce')
NewTrans['Posting Date'] = pd.to_datetime(NewTrans['Posting Date'], errors='coerce')

LastInserted = NewTrans[(NewTrans['Balance'] == LastBalance) & (NewTrans['Posting Date'] == LastPostDate)]

# %%
if LastInserted.empty or NewTrans.empty: 
    sys.exit()
else:
    NewTrans = NewTrans.iloc[:LastInserted.index.min()]    

if NewTrans['Balance'].any() or NewTrans.empty:
    NewBalance = NewTrans[pd.isnull(NewTrans['Balance']) == False].iloc[:1]['Balance'].values[0]
else:
    NewBalance = LastBalance

# %%
LastPostDate

# %%
NewTrans['Description'][pd.isnull(NewTrans['Balance'])] = 'z-Pending' + NewTrans['Description'][pd.isnull(NewTrans['Balance'])]

cols = ['Posting Date', 'Description', 'Amount']
NewTrans = NewTrans[cols]
NewTrans['Description'].replace({'\*':'',"\s+":' '}, regex=True, inplace=True)

display('New Balance: ' + str(NewBalance), 'Downloaded Transactions For Today', NewTrans.head(),'Data Types', NewTrans.dtypes)

# %%
NewTrans

# %% [markdown]
# ## <ins> New Transactions Table Part 2

# %% [markdown]
# <font color='blue'>  
# **Adding in Pay Period, Category, and Sub-Cateogry  
#    Read in Auto tag using active sheet to identify workbook and corresponding dictionary structure  
#   Adding in Category and Sub-Category using dictionary structure  
#   Separate Expenses versus income  
#     Order Columns  
#     Display Columns** 
# <font color='black'>
# 

# %%
# NewTrans
# Temp
# NewTrans.loc[Temp.index.to_list(), col]
# catch
# dct[catch][col]
# dct
# E_Pending_dict
# Autotag_dict

# %%
display('<-------------------------Importing and Scubbing New Transactions Part 2------------------------->')
NewTrans['Pay Period'] = NewTrans['Posting Date'].apply(Compute_PayPeriod)
NewTrans['Category'] = np.nan
NewTrans['Sub-Category'] = np.nan

shtTag = xw.sheets('Autotag')
Autotag = shtTag.range('a1').expand().options(pd.DataFrame, header=1, index=1).value
Autotag_dict = Autotag.to_dict('index')


    
for dct in [E_Pending_dict, N_Pending_dict]:
    for catch in dct:
        Temp = NewTrans.loc[:,['Category', 'Sub-Category', 'Pay Period']][(NewTrans['Description'].str.lower().str.contains(dct[catch]['Description'].lower())) & (pd.isnull(NewTrans['Category'])) & (NewTrans['Amount'] == dct[catch]['Amount'])]
        if not Temp.empty:
            for col in ['Category', 'Sub-Category', 'Pay Period']:
                try:
                    NewTrans.loc[Temp.index.to_list(), col] = dct[catch][col]                        
                except:
                    pass

for col in ['Category', 'Sub-Category']:
    for catch in Autotag_dict.keys():
        NewTrans.loc[:,col][(NewTrans['Description'].str.lower().str.contains(catch.lower())) & (pd.isnull(NewTrans[col]))] = Autotag_dict[catch][col]    

NewTrans['Description'].replace({'\*':'',"\s+":' '}, regex = True, inplace=True)

NewTrans['Description'][pd.isnull(NewTrans['Category']) == False] += ' (A)'
NewTransExpse = NewTrans[NewTrans['Amount'] <= 0]
NewTransNcome = NewTrans[NewTrans['Amount'] > 0]

cols = ['Pay Period', 'Posting Date', 'Description', 'Amount', 'Category', 'Sub-Category']
cols2 = ['Pay Period', 'Posting Date', 'Description', 'Amount', 'Category']
NewTransExpse = NewTransExpse[cols]
NewTransNcome = NewTransNcome[cols2]

display('New Expense Trans READY, ' + str(len(NewTransExpse)) + ' NewRecords',NewTransExpse.head())
display('New Ncome Trans READY, ' + str(len(NewTransNcome)) + ' NewRecords',NewTransNcome.head())

# %%
# NewTransExpse

# %% [markdown]
# ## <ins> Updating Current Transaction With New Transactions

# %% [markdown]
# <font color='blue'>  
# **Identify Space For New Records  
#     Delete old records, Add Space, Insert New Records**  
# <font color='blue'>
# **Update Current Balance**

# %%
if NewTrans.empty:
    display('<-----------------------------------NO NEW TRANSACTIONS----------------------------------------->')
else:
#     Update_Current_Transaction_With_New()
    display('<-------------------------Updating Current Transactions With New Transactions------------------------->')
    NewRecords = shtTrans.range(NcomeTbl(2,1).address,NcomeTbl(len(NewTransNcome)+1,len(NewTransNcome.columns)).address)

    for tbl, dList, New in zip([ExpenseTbl, NcomeTbl], [eList, nList], [NewTransExpse, NewTransNcome]):
        if len(dList) != 0:
            shtTrans.range(tbl(2,1).address,tbl(len(dList)+1,len(New.columns)).address).delete(shift='up')

        if len(New) != 0:
            NewRecords = shtTrans.range(tbl(2,1).address,tbl(len(New)+1,len(New.columns)).address)
            NewRecords.insert(shift = 'down', copy_origin='format_from_right_or_below')

            NewRecords = shtTrans.range(tbl(1,1).address,tbl(len(New)+1,len(New.columns)).address)
            shtTrans.range(tbl(len(New)+2,1).address, tbl(len(New)+2,len(New.columns)).address).color = (169, 208, 142)
            NewRecords.color = None
            NewRecords.options(pd.DataFrame,header=1, index=False).value = New

# %%
shtTrans.range('CBalance').value = NewBalance

# %%
# Autotag_dict
# Autotag

# %%


# %%



