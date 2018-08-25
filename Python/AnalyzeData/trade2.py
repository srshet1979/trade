import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from pandas import DataFrame
from typing import Set, Any
from datetime import datetime
from functools import reduce
from openpyxl import Workbook
import time

#method to delete the columns other than the listed ones
def remove_others(df: DataFrame, columns: Set[Any]):
    cols_total: Set[Any] = set(df.columns)
    diff: Set[Any] = cols_total - columns
    df.drop(diff, axis=1, inplace=True)

def clean_float(x):
    x = x.replace("$", "").replace(",", "").replace(" ", "")
    return float(x)

def clean_int(x):
    x = x.replace("$", "").replace(",", "").replace(" ", "")
    return int(x)

def date(datestr="", format="%d-%m-%Y"):
    from datetime import datetime
    if not datestr:
        return datetime.today().date()
    return datetime.strptime(datestr, format).date()


datestr = open('D:\\trade\downloads\\today\\date').readline()
datefmt=date(datestr.rstrip(),"%d-%m-%Y")
exp_month = datefmt.strftime('%b-%y')



#locad all the data files
data = pd.read_csv('D:\\trade\downloads\\today\\data.csv')

cm=pd.read_csv('D:\\trade\downloads\\today\\cm.csv')
fo=pd.read_csv('D:\\trade\downloads\\today\\fo.csv')
mto=pd.read_csv('D:\\trade\downloads\\today\\mto.csv')
sa=pd.read_excel('D:\\trade\downloads\\SA.xlsx', sheet_name='SA')
#remove unwanted columns and rename the required ones
remove_others(data,{"Symbol","Change","Traded Value(crs)"})
data.columns=['SYMBOL','P2','TV3']

remove_others(cm,{"SYMBOL","TOTALTRADES"})
cm.columns=['SYMBOL','NT3']
remove_others(mto,{'SYMBOL','DQ3'})

# #format data
if (data.P2.dtype == 'object'):
  data['P2']=data['P2'].apply(clean_float)

if (data.TV3.dtype == 'object'):
  data['TV3']=data['TV3'].apply(clean_float)



#filter fo sheet
#filter INSTRUMENT
fo_filter=fo[fo['INSTRUMENT']=='FUTSTK']
#filter expiry month
fo_filter['EXPIRY_DT'] = pd.to_datetime(fo_filter['EXPIRY_DT'])
fo_filter['EXPIRY_DT']=fo_filter['EXPIRY_DT'].apply(lambda x: x.strftime('%b-%y'))
# CurrentMonth = datetime.today().strftime('%b-%y')
filter=fo_filter['EXPIRY_DT'] <= exp_month
fo=fo_filter.loc[filter]
#truncate other columns
remove_others(fo,{"SYMBOL","CONTRACTS","CHG_IN_OI"})
fo.columns=['SYMBOL','COI2','OIC3']

#data,cm,mto,fo

data_cm=pd.merge(data,cm,on='SYMBOL', left_index=True,right_index=True,how='inner')
data_cm_mto=pd.merge(data_cm,mto,on='SYMBOL', left_index=True,right_index=True,how='inner')
data_cm_mto_fo=pd.merge(data_cm_mto,fo,on='SYMBOL', left_index=True,right_index=False,how='inner')

#prepare master spreadsheet to update
cols = sa.columns.tolist()
sa['P1']=sa['P2']
sa = sa.drop('P2', 1)

sa['NT1']=sa['NT2']
sa['NT2']=sa['NT3']
sa = sa.drop('NT3', 1)

sa['TV1']=sa['TV2']
sa['TV2']=sa['TV3']
sa = sa.drop('TV3', 1)

sa['DQ1']=sa['DQ2']
sa['DQ2']=sa['DQ3']
sa = sa.drop('DQ3', 1)

sa['COI1']=sa['COI2']
sa = sa.drop('COI2', 1)

sa['OIC1']=sa['OIC2']
sa['OIC2']=sa['OIC3']
sa = sa.drop('OIC3', 1)

#print(sa)
#sa=pd.merge(sa, data_cm_mto_fo, how='inner', on='SYMBOL',left_index=False, right_index=False)


sa=pd.merge(sa, data_cm_mto_fo, on='SYMBOL', left_index=True,right_index=False,how='inner')
sa = sa[cols]
sa.sort_values('SYMBOL')

#apply formulas

sa.NTC1=sa.NT2-sa.NT1
sa.NTC2=sa.NT3-sa.NT2
sa.NTPC1=(sa.NTC1/sa.NT1)*100
sa.NTPC2=(sa.NTC2/sa.NT2)*100

sa.TVC1=sa.TV2-sa.TV1
sa.TVC2=sa.TV3-sa.TV2

sa.TVPC1=(sa.TVC1/sa.TV1)*100
sa.TVPC2=(sa.TVC2/sa.TV2)*100

sa.DQC1=sa.DQ2-sa.DQ1
sa.DQC2=sa.DQ3-sa.DQ2
sa.DQPC1=(sa.DQC1/sa.DQ1)*100
sa.DQPC2=(sa.DQC2/sa.DQ2)*100

sa.OICC1=sa.OIC2-sa.OIC1
sa.OICC2=sa.OIC3-sa.OIC2
sa.OICPC1=(sa.OICC1/sa.OIC1)*100
sa.OICPC2=(sa.OICC2/sa.OIC2)*100

#export to xl
writer = pd.ExcelWriter('D:\\trade\downloads\\today\\SA.xlsx')
sa.to_excel(writer,'SA')
writer.save()

writer = pd.ExcelWriter('D:\\trade\downloads\\SA.xlsx')
sa.to_excel(writer,'SA')
writer.save()
