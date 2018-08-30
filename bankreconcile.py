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
DateFrom = "{d'2018-01-01'}"
DateTo = "{d'2018-08-31'}"

# open ODBC connection to Quickbooks and run sp_report to query UnCleared Credit Card Transaction
cn = pyodbc.connect('DSN=QuickBooks Data;')

sql = """sp_report CustomTxnDetail show Date, Account, TxnType, ClearedStatus, Debit, Credit
parameters DateFrom = """+ DateFrom +""", DateTo = """+ DateTo +""", SummarizeRowsBy = 'TotalOnly', AccountFilterType = 'Bank'
where RowType = 'DataRow' and AccountFullName Like '%A-Woodforest LLC 3221%' and ClearedStatus <> 'Cleared'
ORDER BY Credit ASC"""

#load data to DataFrame2
df2 = pd.read_sql(sql,cn)

df2['Debit'] = df2['Debit'].replace(np.nan,0)
df2['Credit'] = df2['Credit'].replace(np.nan,0)

df2['Transaction_Amount'] = df2['Credit'] - df2['Debit']

df2.drop(['ClearedStatus','Debit','Credit',], axis=1,inplace=True)

df2 = df2.sort_values(['Date','Transaction_Amount'],ascending=[True,True])

df2 = df2[df2['Transaction_Amount'] != 0]

print (df2)