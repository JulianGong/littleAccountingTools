# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 13:02:24 2017

@author: 028375
"""

from __future__ import unicode_literals, division
import pandas as pd
import os.path
import numpy as np
import copy



def check1(collateral):
    tmp=['产品编号','产品名称','发行部门','交易类别','成交份额','成交金额','手续费(收取)','成本','交易损益']
    collateral=copy.deepcopy(collateral[tmp])
    
    collateral['手续费(收取)']=pd.to_numeric(collateral['手续费(收取)'])
    collateral['成本']=pd.to_numeric(collateral['成本'])
    collateral['交易损益']=pd.to_numeric(collateral['交易损益'])
    collateral['交易损益']=collateral['交易损益'].replace(np.nan,0.)
    
    collateral['交易类别标示']=0
    tmp5=['交易类别标示']
    
    collateral.loc[collateral['交易类别']=='认购',tmp5]=1
    collateral.loc[collateral['交易类别']=='清盘',tmp5]=2
    collateral.loc[collateral['交易类别']=='强赎',tmp5]=2
    collateral.loc[collateral['交易类别']=='赎回',tmp5]=2
    collateral.loc[collateral['交易类别']=='分红',tmp5]=3
    
    tmp2=[['股衍部',138],['证金部',146],['大宗部',138]]
    tmp2=pd.DataFrame(tmp2,columns=['发行部门','成本中心段'])
    collateral=pd.merge(collateral,tmp2,how='left',on='发行部门')
    
    collateral_drop=collateral[collateral['交易类别标示']==0]
    collateral=collateral[collateral['交易类别标示']!=0]
    
    collateral_drop1=collateral[np.isnan(collateral['成本中心段'])]
    collateral_drop=collateral_drop.append(collateral_drop1)
    
    collateral=(collateral[~np.isnan(collateral['成本中心段'])]).reset_index(drop=True)
    collateral['成本中心段']=(collateral['成本中心段']).astype(str)
    
    return collateral,collateral_drop




def genNote(collateral):
    Notes_cols=['公司段','成本中心段','个人段','科目段','子目段','项目段','往来段','产品段','币种段','借方','贷方','摘要']
    
    Notes_temp1=[['0001','000','0000','900006',  '000000','000000','0000','00', '0000','','',''],
                 ['0001','',   '0000','21020100','000000','000000','0000','324','0000','','','']]
    Notes_temp1=pd.DataFrame(Notes_temp1,columns=Notes_cols)
    
    Notes_temp2=[['0001','000','0000','900006',  '000000','000000','0000','00', '0000','','',''],
                 ['0001','',   '0000','21020100','000000','000000','0000','324','0000','','',''],
                 ['0001','',   '0000','61111200','611104','000000','0000','324','0000','','','']]
    Notes_temp2=pd.DataFrame(Notes_temp2,columns=Notes_cols)
    
    Notes_temp3=[['0001','000','0000','900006',  '000000','000000','0000','00', '0000','','',''],
                 ['0001','',   '0000','61111200','611104','000000','0000','324','0000','','','']]
    Notes_temp3=pd.DataFrame(Notes_temp3,columns=Notes_cols)
    
    Notes_temp4=[['0001','',   '0000','60219900','000000','000000', '0000','324', '0000','','',''],
                 ['0001','000','0000','22211202','000000','VAT0028','0000','00',  '0000','','','']]
    Notes_temp4=pd.DataFrame(Notes_temp4,columns=Notes_cols)
    
    
    Notes=pd.DataFrame(np.zeros((0,12)),columns=Notes_cols)
    for i in list(range(len(collateral))):
        
        if collateral['交易类别标示'][i]==1:
            iNotes=copy.deepcopy(Notes_temp1)
            iNotes['借方'][0]=collateral['成交金额'][i]
            iNotes['贷方'][1]=collateral['成本'][i]
            
        elif collateral['交易类别标示'][i]==2:
            iNotes=copy.deepcopy(Notes_temp2)
            iNotes['贷方'][0]=collateral['成交金额'][i]
            iNotes['借方'][1]=collateral['成本'][i]
            if collateral['交易损益'][i]>=0:
                iNotes['借方'][2]=collateral['交易损益'][i]
            else:
                iNotes['贷方'][2]=-collateral['交易损益'][i]
            
        elif collateral['交易类别标示'][i]==3:
            iNotes=copy.deepcopy(Notes_temp3)
            if collateral['交易损益'][i]>=0:
                iNotes['借方'][1]=collateral['交易损益'][i]
                iNotes['贷方'][0]=collateral['交易损益'][i]
            else:
                iNotes['借方'][0]=collateral['交易损益'][i]
                iNotes['贷方'][1]=collateral['交易损益'][i]
                
        iNotes['成本中心段'][1:]=collateral['成本中心段'][i]
        iNotes['摘要']=collateral['产品名称'][i]+collateral['交易类别'][i]+collateral['产品编号'][i]
        
        if collateral['手续费(收取)'][i]>0:
            iNotes_add=copy.deepcopy(Notes_temp4)
            iNotes_add['成本中心段'][0]=collateral['成本中心段'][i]
            iNotes_add['贷方'][0]=(collateral['手续费(收取)'][i]/1.06).round(decimals=2)
            iNotes_add['贷方'][1]=(collateral['手续费(收取)'][i]/1.06*0.06).round(decimals=2)
            iNotes_add['摘要'][0]=collateral['产品名称'][i]+collateral['交易类别'][i]+collateral['产品编号'][i]+'手续费收入'
            iNotes_add['摘要'][1]=collateral['产品名称'][i]+collateral['交易类别'][i]+collateral['产品编号'][i]+'手续费收入销项税'
            iNotes=(iNotes.append(iNotes_add)).reset_index(drop=True)
        
        Notes=Notes.append(iNotes).reset_index(drop=True)
    
    return Notes


if __name__=="__main__":
    path0=os.path.dirname(os.path.realpath(__file__))+'//'
    
    path1='ELN\Collateral2.xlsx'
    path2='ELN\Report2.xlsx'   
    
    collateral=pd.read_excel(path0+path1,0,encoding="gbk",keep_default_na=False)
    
    collateral,collateral_drop=check1(collateral)
    Notes=genNote(collateral)
    
    
    
    