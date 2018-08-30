import pandas as pd
import numpy as np
from pandas import DataFrame
import pyodbc
import os
import sys


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BankStatementPath = os.path.join(BASE_DIR,'BankStatement.xlsx')
OutputExcelPath = os.path.join(BASE_DIR,'Reconciliation.xlsx')

# open Excel file from bank statement, create dataframe from worksheet
df = pd.read_excel(BankStatementPath, header=0, dtype={'Customer Ref':str})
df['Debit'] = df['Debit'].replace(np.nan,0)
df['Credit'] = df['Credit'].replace(np.nan,0)
df['Transaction_Amount'] =  df['Credit'] - df['Debit']


#drop unneccessary columns
df = df.drop(df.columns[[1,3,5,6,7]],axis = 1)

#sort the dataframe by Transaction Amount
df = df.sort_values(['Transaction_Amount','Trans Date'],ascending=[True,True])

list1 = []
counter1 = []

for index, row in df.iterrows():
    list1.append(row['Transaction_Amount'])
    counter1.append(list1.count(row['Transaction_Amount']))

df['Counter'] = counter1
df['Combine'] = df['Transaction_Amount'].astype(str) + '|' + df['Counter'].astype(str)

#print (df)
print (df.dtypes)










#DateFrom and DateTo paramters for the query
DateFrom = "{d'2018-01-01'}"
DateTo = "{d'2018-08-31'}"

# open ODBC connection to Quickbooks and run sp_report to query UnCleared Credit Card Transaction
cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = """sp_report CustomTxnDetail show Date, Account, TxnType, ClearedStatus, Debit, Credit
parameters DateFrom = """+ DateFrom +""", DateTo = """+ DateTo +""", SummarizeRowsBy = 'TotalOnly', AccountFilterType = 'Bank'
where RowType = 'DataRow' and AccountFullName Like '%A-Woodforest LLC 3221%' and ClearedStatus <> 'Cleared'
ORDER BY Credit ASC"""

#load data to DataFrame2
df2 = pd.read_sql(sql,cn, parse_dates=['Date'])

df2['Debit'] = df2['Debit'].replace(np.nan,0)
df2['Credit'] = df2['Credit'].replace(np.nan,0)

df2['Transaction_Amount'] =  df2['Debit'] - df2['Credit']

df2.drop(['ClearedStatus','Debit','Credit',], axis=1,inplace=True)

df2 = df2.sort_values(['Transaction_Amount'],ascending=[True])

df2 = df2[df2['Transaction_Amount'] != 0]

list3 = []
counter2 = []

for index, row in df2.iterrows():
    list3.append(row['Transaction_Amount'])
    counter2.append(list3.count(row['Transaction_Amount']))

df2['Counter'] = counter2    
df2['Combine'] = df2['Transaction_Amount'].astype(str) + '|' + df2['Counter'].astype(str)

print (df2.dtypes)

df['Matched'] = df['Combine'].isin(df2['Combine'])
df2['Matched'] = df2['Combine'].isin(df['Combine'])


writer = pd.ExcelWriter(OutputExcelPath,engine='xlsxwriter')
df.to_excel(writer,sheet_name='Sheet1',startcol=0,startrow=0,index=False,header=True,engine='xlsxwriter')
df2.to_excel(writer,sheet_name='Sheet1',startcol=10,startrow=0,index=False,header=True,engine='xlsxwriter')


writer.save()
cn.close()