import pandas as pd
import numpy as np
from pandas import DataFrame
import pyodbc
import os
import sys


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BankStatementPath = os.path.join(BASE_DIR,'BankStatement.xlsx')
OutputExcelPath = os.path.join(BASE_DIR,'Reconciliation.xlsx')

#DateFrom and DateTo paramters for the query 
DateFrom = "{d'2019-01-01'}"
DateTo = "{d'2019-08-09'}"

# open Excel file from bank statement, create dataframe from worksheet
df = pd.read_excel(BankStatementPath, header=0, dtype={'Reference':str})
df['Debit Amount'] = df['Debit Amount'].replace(np.nan,0)
df['Credit Amount'] = df['Credit Amount'].replace(np.nan,0)

def getpendingcheckno(row):
    if "Pending Check" in row['Memo']:
        checkno = row['Memo'][-5:]
    else:
        checkno = row['Reference']
    return checkno

df['Reference'] = df.apply(getpendingcheckno, axis = 1)

#drop unneccessary columns
df = df.drop(columns = ['Record Type','Account Number', 'Account Name','Code'],axis = 1)
df.rename(columns={'Credit Amount':'Credit','Debit Amount':'Debit'},inplace = True)

#sort the dataframe by Transaction Amount

Debit = df[df['Debit'] != 0]
Debit = Debit.drop(columns = ['Credit'])
Debit = Debit.sort_values(['Debit','Date'], ascending = [True, True])



Credit = df[df['Credit'] !=0]
Credit = Credit.drop(columns = ['Debit'])
Credit = Credit.sort_values(['Credit','Date'],ascending = [True,True])

Check = Debit[Debit['Reference'].str.contains('^\d{5}$',regex = True)]
Debit = Debit.merge(Check,how = 'left', indicator = True)
OtherDebit = Debit[Debit['_merge'] == 'left_only']
OtherDebit.drop(['_merge'],axis = 1, inplace = True)

list1 = []
counter1 = []

for index, row in OtherDebit.iterrows():
    list1.append(row['Debit'])
    counter1.append(list1.count(row['Debit']))


OtherDebit.loc[:,'Counter'] = counter1


OtherDebit.loc[:,'Combine'] = OtherDebit['Debit'].astype(str) + '|' + OtherDebit['Counter'].astype(str)



list3 = []
counter3 = []

for index, row in Credit.iterrows():
    list3.append(row['Credit'])
    counter3.append(list3.count(row['Credit']))

Credit.loc[:,'Counter'] = counter3
Credit.loc[:,'Combine'] = Credit['Credit'].astype(str) + '|' + Credit['Counter'].astype(str)



# create a new comlum to combine transaction amount and check number
Check.loc[:,'Combine'] = Check['Debit'].astype(str) + '|' + Check['Reference'].astype(str)




# open ODBC connection to Quickbooks and run sp_report to query UnCleared Credit Card Transaction
cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = """sp_report CustomTxnDetail show Date, Account, TxnType , RefNumber, ClearedStatus, Debit, Credit
parameters DateFrom = """+ DateFrom +""",DateTo = """+ DateTo +""", SummarizeRowsBy = 'TotalOnly', AccountFilterType = 'Bank'
where RowType = 'DataRow' and AccountFullName Like '%A-Woodforest LLC 3221%' and ClearedStatus <> 'Cleared'
ORDER BY Credit ASC"""

#load data to DataFrame2
df2 = pd.read_sql(sql,cn, parse_dates=['Date'])

df2['Debit'] = df2['Debit'].replace(np.nan,0)
df2['Credit'] = df2['Credit'].replace(np.nan,0)
df2['RefNumber'] = df2['RefNumber'].replace(np.nan,0)
df2.rename(columns = {'Debit':'Credit','Credit':'Debit'}, inplace = True)


df2.drop(['ClearedStatus',], axis=1,inplace=True)

# remove rows that have transaction amount = 0


Debit2 = df2[df2['Debit'] != 0]
Debit2.drop(['Credit',], axis = 1, inplace = True)
Debit2 = Debit2.sort_values(['Debit','Date'], ascending = [True, True])

Credit2 = df2[df2['Credit'] != 0]
Credit2.drop(['Debit',], axis = 1, inplace = True)
Credit2 = Credit2.sort_values(['Credit','Date'],ascending = [True, True])


# use regular expression to find check transactions
Check2 = Debit2[Debit2['RefNumber'].str.contains('^\d{5}$',regex = True, na=False)]

#this filter is to remove NSF Checks from Customer out of regular checks payment
Check2 = Check2[Check2['TxnType'] != 'Invoice']

#combine Check amount with Check Number
Check2.loc[:,'Combine'] = Check2['Debit'].astype(str) + '|' + Check2['RefNumber'].astype(str)

Debit2 = Debit2.merge(Check2,how = 'left', indicator = True)
OtherDebit2 = Debit2[Debit2['_merge'] == 'left_only']
OtherDebit2.drop(['_merge'],axis = 1, inplace = True)

list2 = []
counter2 = []

for index, row in OtherDebit2.iterrows():
    list2.append(row['Debit'])
    counter2.append(list2.count(row['Debit']))


OtherDebit2.loc[:,'Counter'] = counter2


OtherDebit2.loc[:,'Combine'] = OtherDebit2['Debit'].astype(str) + '|' + OtherDebit2['Counter'].astype(str)


list4 = []
counter4 = []

for index, row in Credit2.iterrows():
    list4.append(row['Credit'])
    counter4.append(list4.count(row['Credit']))

Credit2.loc[:,'Counter'] = counter4
Credit2.loc[:,'Combine'] = Credit2['Credit'].astype(str) + '|' + Credit2['Counter'].astype(str)



Check.loc[:,'Matched'] = Check['Combine'].isin(Check2['Combine'])
Check2.loc[:,'Matched'] = Check2['Combine'].isin(Check['Combine'])


OtherDebit.loc[:,'Matched'] = OtherDebit['Combine'].isin(OtherDebit2['Combine'])
OtherDebit2.loc[:,'Matched'] = OtherDebit2['Combine'].isin(OtherDebit['Combine'])


Credit.loc[:,'Matched'] = Credit['Combine'].isin(Credit2['Combine'])
Credit2.loc[:,'Matched'] = Credit2['Combine'].isin(Credit['Combine'])


writer = pd.ExcelWriter(OutputExcelPath,engine='xlsxwriter')
numberformat = writer.book.add_format({'num_format': '#,##0.00'})


Check.to_excel(writer,sheet_name='Checks',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')
Check2.to_excel(writer,sheet_name='Checks',startcol=10,startrow=0,index=False,header=True,engine='xlsxwriter')
writer.sheets['Checks'].set_column('B:B', None, numberformat)
writer.sheets['Checks'].set_column('O:O', None, numberformat)
writer.sheets['Checks'].autofilter('B1:Q1')



OtherDebit.to_excel(writer,sheet_name='OtherDebits', startcol=0, startrow = 0, index = False, header = True, engine = 'xlsxwriter')
OtherDebit2.to_excel(writer,sheet_name='OtherDebits', startcol=10, startrow = 0, index = False, header = True, engine = 'xlsxwriter')
writer.sheets['OtherDebits'].set_column('B:B', None, numberformat)
writer.sheets['OtherDebits'].set_column('O:O', None, numberformat)
writer.sheets['OtherDebits'].autofilter('B1:R1')


Credit.to_excel(writer,sheet_name='Credits',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')
Credit2.to_excel(writer,sheet_name='Credits',startcol=10,startrow=0,index=False,header=True,engine='xlsxwriter')
writer.sheets['Credits'].set_column('B:B', None, numberformat)
writer.sheets['Credits'].set_column('O:O', None, numberformat)
writer.sheets['Credits'].autofilter('B1:R1')


writer.save()
cn.close()


#automatically open the Reconciliation.xls from Excel
os.startfile(OutputExcelPath)
