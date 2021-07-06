# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 14:15:52 2017

@author: 028375
"""

from __future__ import unicode_literals, division
import pandas as pd
import os.path
import numpy as np
import copy


def check1(collateral):
    
    collateral['现金流产生日期']=collateral['现金流产生日期'].astype(str)
    collateral['交易对手编号']=collateral['交易对手编号'].astype(str)
    collateral['确认金额(结算货币)']=pd.to_numeric(collateral['确认金额(结算货币)'])
    
    tmp_col1=['提取预付金','提交预付金','内转']
    tmp_col1=pd.DataFrame(tmp_col1,columns=['现金流类型'])
    tmp_col1['现金流类型标示']=1
    
    tmp_col2=['前端支付','前端期权费','展期期权费','到期结算','部分赎回','全部赎回','期间结算','红利支付','定结期间结算']
    tmp_col2=pd.DataFrame(tmp_col2,columns=['现金流类型'])
    tmp_col2['现金流类型标示']=2
    
    tmp_col=tmp_col1.append(tmp_col2).reset_index(drop=True)
    
    collateral=pd.merge(collateral,tmp_col,how='left',on='现金流类型')
    
    collateral_drop=collateral[np.isnan(collateral['现金流类型标示'])]
    collateral=(collateral[~np.isnan(collateral['现金流类型标示'])]).reset_index(drop=True)
    
    return collateral, collateral_drop



def genNotes(collateral):
    
    tmp=['交易对手编号','交易编号','现金流类型','现金流产生日期','确认金额(结算货币)','结算方式','往来段','现金流类型标示']
    collateral=collateral[tmp]
    
    Notes_cols=['公司段','成本中心段','个人段','科目段','子目段','项目段','往来段','产品段','币种段','借方','贷方','摘要']
    
    Notes_temp1=[['0001','000','0000','900006',  '000000','000000','0000','00','0000','','',''],
                 ['0001','000','0000','22023700','000000','000000', '',   '00','0000','','','']]
    
    Notes_temp1=pd.DataFrame(Notes_temp1,columns=Notes_cols)
    
    Notes_temp2=[['0001','000','0000','900006',  '000000','000000','0000',  '00', '0000','','',''],
                 ['0001','138','0000','61110700','611104','000000','610000','312','0000','','','']]
    
    Notes_temp2=pd.DataFrame(Notes_temp2,columns=Notes_cols)
    
    Notes_temp3=[['0001','000','0000','22023700','000000','000000','',      '00', '0000','','',''],
                 ['0001','138','0000','61110700','611104','000000','610000','312','0000','','','']]
    
    Notes_temp3=pd.DataFrame(Notes_temp3,columns=Notes_cols)
    
    
    Notes=pd.DataFrame(np.zeros((0,12)),columns=Notes_cols)
    
    for i in list(range(len(collateral))):
        
        if collateral['现金流类型标示'][i]==1:
            iNotes=copy.deepcopy(Notes_temp1)
            iNotes['往来段'][1]=collateral['往来段'][i]
            iNotes['摘要'][0]=collateral['现金流产生日期'][i]+collateral['现金流类型'][i]+collateral['交易对手编号'][i]
        elif (collateral['现金流类型标示'][i]==2)&(collateral['结算方式'][i]=='现金'):
            iNotes=copy.deepcopy(Notes_temp2)
            iNotes['摘要'][0]=collateral['现金流产生日期'][i]+collateral['现金流类型'][i]+collateral['交易编号'][i]
        elif (collateral['现金流类型标示'][i]==2)&(collateral['结算方式'][i]=='预付金'):
            iNotes=copy.deepcopy(Notes_temp3)
            iNotes['往来段'][0]=collateral['往来段'][i]
            iNotes['摘要'][0]=collateral['现金流产生日期'][i]+collateral['现金流类型'][i]+collateral['交易编号'][i]
        
        if collateral['确认金额(结算货币)'][i]>=0:
            iNotes['借方'][0]=collateral['确认金额(结算货币)'][i]
            iNotes['贷方'][1]=collateral['确认金额(结算货币)'][i]
        else:
            iNotes['借方'][1]=-collateral['确认金额(结算货币)'][i]
            iNotes['贷方'][0]=-collateral['确认金额(结算货币)'][i]
            
        iNotes['摘要'][1]=iNotes['摘要'][0]
        Notes=Notes.append(iNotes).reset_index(drop=True)
    
    return Notes
    



if __name__=="__main__":
    path0=os.path.dirname(os.path.realpath(__file__))+'//'
    
    path1='Opt_DM\Collateral2.xlsx'
    path2='Opt_DM\Report2.xlsx'   
    
    collateral=pd.read_excel(path0+path1,0,encoding="gbk",keep_default_na=False)
    
    collateral, collateral_drop=check1(collateral)
    
    Notes=genNotes(collateral)

        