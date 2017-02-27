

import pandas as pd
import datetime
import xlsxwriter
import numpy as np

Now = datetime.datetime.now()

ql_1 = pd.read_excel('Z:\\Inventory position reporting\\Quotelog Data\\Quotelog.xlsx')
ql_1['Age'] = (ql_1['today'] - ql_1['date Scheduled']).astype('timedelta64[D]')

ql_2_1 = ql_1[ql_1['Customer'].str.contains('Comcast',na=False)].copy()
ql_2_1['CustomerType'] = 1
ql_1 = ql_1.append(ql_2_1.copy())
ql_1.drop_duplicates(subset = 'SO',keep=False, inplace=True)
ql_2_2 = ql_1[ql_1['Customer'].str.contains('Charter',na=False)].copy()
ql_2_2['CustomerType'] = 1
ql_1 = ql_1.append(ql_2_2.copy())
ql_1.drop_duplicates(subset = 'SO',keep=False,inplace=True)
ql_2_3 = ql_1[ql_1['Customer'].str.contains('Cox',na=False)].copy()
ql_2_3['CustomerType'] = 1
ql_1 = ql_1.append(ql_2_3.copy())
ql_1.drop_duplicates(subset = 'SO',keep=False,inplace=True)
ql_2_4 = ql_1.copy()
ql_2_4['CustomerType'] = 2

ql_3_1 = ql_2_1.copy()
ql_3_1 = ql_3_1.append(ql_2_2.copy())
ql_3_1 = ql_3_1.append(ql_2_3.copy())
ql_3_1 = ql_3_1.append(ql_2_4.copy())
ql_3_1['Net Value'] = 'NaN'

ql_3_3 = ql_3_1[ql_3_1['CustomerType']==2].copy()
ql_3_3_1 = ql_3_3[ql_3_3['Age'] <= 10].copy()
ql_3_3_1['Net Value'] = ql_3_3_1['Amt'] * 0.9
ql_3_3 = ql_3_3.append(ql_3_3_1.copy())
ql_3_3.drop_duplicates(subset = 'SO', keep=False, inplace = True)
ql_3_3_2 = ql_3_3.copy()
ql_3_3_2['Net Value'] = 0
ql_3_1 = ql_3_1.append(ql_3_3_1.copy())
ql_3_1 = ql_3_1.append(ql_3_3_2.copy())
ql_3_1.drop_duplicates(subset = 'SO',keep=False,inplace=True)
ql_4_1 = ql_3_3_1.copy()
ql_4_1 = ql_4_1.append(ql_3_3_2.copy())

ql_3_2_1 = ql_3_1[['Age','Amt','Customer','CustomerType','Odds of Job','SO','SalesPerson','date Scheduled','date created','today', 'Net Value', 'Last Contact', 'Quote Status','lines']][ql_3_1['Odds of Job'].notnull()].copy()
ql_3_2_1['Net Value'] = ql_3_2_1['Amt'] * ql_3_2_1['Odds of Job']
ql_4_1 = ql_4_1.append(ql_3_2_1.copy())
ql_3_1_1 = ql_3_1[['Age','Amt','Customer','CustomerType','Odds of Job','SO','SalesPerson','date Scheduled','date created','today', 'Net Value','Last Contact', 'Quote Status','lines']][ql_3_1['Odds of Job'].isnull()].copy()
ql_3_1_1_1 = ql_3_1_1.loc[ql_3_1_1['Age'] <= 10].copy()
ql_3_1_1_1['Net Value'] = ql_3_1_1_1['Amt'] * .5
ql_4_1 = ql_4_1.append(ql_3_1_1_1.copy())
ql_3_1_1_2 = ql_3_1_1.loc[(ql_3_1_1['Age'] > 10) & (ql_3_1_1['Age'] <= 30)].copy()
ql_3_1_1_2['Net Value'] = ql_3_1_1_2['Amt'] * .25
ql_4_1 = ql_4_1.append(ql_3_1_1_2.copy())
ql_3_1_1_3 = ql_3_1_1.loc[(ql_3_1_1['Age'] > 30) & (ql_3_1_1['Age'] <= 60)].copy()
ql_3_1_1_3['Net Value'] = ql_3_1_1_3['Amt'] * .05
ql_4_1 = ql_4_1.append(ql_3_1_1_3.copy())
ql_3_1_1_4 = ql_3_1_1.loc[ql_3_1_1['Age'] > 60].copy()
ql_3_1_1_4['Net Value'] = ql_3_1_1_4['Amt'] * .025
ql_4_1 = ql_4_1.append(ql_3_1_1_4.copy())

ql_4_1 = ql_4_1[['SalesPerson','SO','Customer','date Scheduled','Age','Amt','lines','Last Contact','Odds of Job','Quote Status','Net Value']]
writer = pd.ExcelWriter('Z:\Inventory position reporting\Quotelog Data\Quotelognew.xlsx',engine='xlsxwriter')
ql_4_1.to_excel(writer,sheet_name='RawData',index=False)
writer.save()




