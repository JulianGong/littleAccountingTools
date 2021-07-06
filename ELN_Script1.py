# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 10:57:19 2017

@author: 028375
"""


from __future__ import unicode_literals, division
import pandas as pd
import os.path
import numpy as np



def check1(lastmonth,thismonth,collateral,lastdate,spotdate):
    
    collateral['日期']=pd.to_datetime(collateral['日期'])
    collateral=collateral[collateral['日期']>lastdate]
    collateral=collateral[collateral['日期']<=spotdate]
    
    lastmonth_dupl=lastmonth[lastmonth.duplicated(subset='产品编号')]
    thismonth_dupl=thismonth[thismonth.duplicated(subset='产品编号')]
    collateral_dupl=collateral[collateral.duplicated()]
    
    lastmonth=lastmonth.drop_duplicates(subset='产品编号')
    thismonth=thismonth.drop_duplicates(subset='产品编号')
    collateral=collateral.drop_duplicates()
    
    lastmonth['客户持有份额']=pd.to_numeric(lastmonth['客户持有份额'])
    lastmonth['成本']=pd.to_numeric(lastmonth['成本'])
    
    thismonth['客户持有份额']=pd.to_numeric(thismonth['客户持有份额'])
    thismonth['成本']=pd.to_numeric(thismonth['成本'])
    
    collateral['成交份额']=pd.to_numeric(collateral['成交份额'])
    
    return lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl
    

def check2(lastmonth,thismonth,collateral):
    
    ContractID=(thismonth['产品编号'].append(lastmonth['产品编号'])).append(collateral['产品编号']).drop_duplicates()
    
    Outputs=pd.DataFrame(ContractID).reset_index(drop=True)
    
    lastmonth=lastmonth[['产品编号','客户持有份额','成本']]
    lastmonth=lastmonth.rename(columns={'客户持有份额':'客户持有份额_期初表', '成本':'成本_期初表'})
    Outputs=pd.merge(Outputs,lastmonth,how='left',on='产品编号')
    
    thismonth=thismonth[['产品编号','客户持有份额','成本']]
    thismonth=thismonth.rename(columns={'客户持有份额':'客户持有份额_期末表', '成本':'成本_期末表'})
    Outputs=pd.merge(Outputs,thismonth,how='left',on='产品编号')
    
    collateral=collateral[['产品编号','交易类别','成交份额']]
    
    tmp1=collateral[collateral['交易类别']=='认购'][['产品编号','成交份额']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='产品编号')
    Outputs=Outputs.rename(columns={'成交份额':'成交份额_认购'})
    
    tmp2=collateral[(collateral['交易类别']=='强赎')|(collateral['交易类别']=='清盘')][['产品编号','成交份额']]
    Outputs=pd.merge(Outputs,tmp2,how='left',on='产品编号')
    Outputs=Outputs.rename(columns={'成交份额':'成交份额_强赎/清盘'})
    
    tmp3=collateral[collateral['交易类别']=='赎回'][['产品编号','成交份额']]
    tmp3=tmp3.groupby(by=['产品编号'],as_index=False)['成交份额'].sum()
    Outputs=pd.merge(Outputs,tmp3,how='left',on='产品编号')
    Outputs=Outputs.rename(columns={'成交份额':'成交份额_赎回'})
    
    Outputs=Outputs.replace(np.nan,0.)
    Outputs['Status']=''
    
    tmp4=Outputs['客户持有份额_期末表']-Outputs['客户持有份额_期初表']-Outputs['成交份额_认购']+Outputs['成交份额_强赎/清盘']+Outputs['成交份额_赎回']
    
    Outputs.loc[tmp4.round(decimals=1)==0,['Status']]='正常'
    Outputs.loc[tmp4.round(decimals=1)!=0,['Status']]='异常'
    
    return Outputs


if __name__=="__main__":
    path0=os.path.dirname(os.path.realpath(__file__))+'//'
    
    path1='ELN\TheBegin.xlsx'
    path2='ELN\TheEnd.xlsx'
    path3='ELN\Collateral.xlsx'
    path4='ELN\Report.xlsx'   
    
    spotdate=pd.to_datetime('2018-03-31')
    lastdate=pd.to_datetime('2018-02-28')
    
    lastmonth=pd.read_excel(path0+path1,0,encoding="gbk",keep_default_na=False)
    thismonth=pd.read_excel(path0+path2,0,encoding="gbk",keep_default_na=False)
    collateral=pd.read_excel(path0+path3,0,encoding="gbk",keep_default_na=False)
    
    lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl=check1(lastmonth,thismonth,collateral,lastdate,spotdate)
    Outputs=check2(lastmonth,thismonth,collateral)
    