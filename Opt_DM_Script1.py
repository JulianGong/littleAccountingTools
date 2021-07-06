# -*- coding: utf-8 -*-
"""
Created on Tue Nov 7 15:14:59 2017

@author: 028375
"""

from __future__ import unicode_literals, division
import pandas as pd
import os.path
import numpy as np


def get_Cost(Outputs,DataFrame0,DateType0,Date0,CostType0,flag0):
    if flag0==1:
        Cost0=DataFrame0[DataFrame0[DateType0]<=Date0][['ContractID','Upfront结算货币']]
    elif flag0==0:
        Cost0=DataFrame0[DataFrame0[DateType0]>Date0][['ContractID','Upfront结算货币']]
    
    Cost0['Upfront结算货币']=-Cost0['Upfront结算货币']
    Outputs=pd.merge(Outputs,Cost0,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Upfront结算货币':CostType0})
    Outputs[CostType0]=Outputs[CostType0].replace(np.nan,0.)
    return Outputs


def get_Cost_NPV(Outputs,DataFrame0,DateType0,Date0,CostType1,CostType2,CostType3,NPVType):
    Outputs=get_Cost(Outputs,DataFrame0,DateType0,Date0,CostType1,1)
    Outputs=get_Cost(Outputs,DataFrame0,DateType0,Date0,CostType2,0)
    Outputs[CostType3]=Outputs[CostType1]+Outputs[CostType2]
    
    NPV0=DataFrame0[['ContractID','RMB合约估值']]
    Outputs=pd.merge(Outputs,NPV0,how='left',on='ContractID')
    
    Outputs=Outputs.rename(columns={'RMB合约估值':NPVType})
    Outputs[NPVType]=Outputs[NPVType].replace(np.nan,0.)    
    return Outputs    



def Collateral_Separate(Outputs,collateral,CollateralType):
    tmp=collateral[collateral['现金流类型']==CollateralType][['ContractID','确认金额(结算货币)']]
    tmp=tmp.groupby(by=['ContractID'],as_index=False)['确认金额(结算货币)'].sum()
    
    Outputs=pd.merge(Outputs,tmp,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'确认金额(结算货币)':CollateralType})
    Outputs[CollateralType]=Outputs[CollateralType].replace(np.nan,0.)
    return Outputs



def Check2(lastmonth,thismonth,collateral,lastdate,spotdate):
    #成本&估值
    ContractID=thismonth['ContractID'].append(lastmonth['ContractID']).drop_duplicates()
    Outputs=pd.DataFrame(ContractID).reset_index(drop=True)
    
    Cost00='成本_期初表_到期'
    Cost01='成本_期初表_存续'
    Cost0='成本_期初表'
    NPV0='估值_期初表'
    Cost10='成本_期末表_存续'
    Cost11='成本_期末表_新起'
    Cost1='成本_期末表'
    NPV1='估值_期末表'
    Outputs=get_Cost_NPV(Outputs,lastmonth,'MATURITYDATEREAL',spotdate,Cost00,Cost01,Cost0,NPV0)
    Outputs=get_Cost_NPV(Outputs,thismonth,'EffectDate',lastdate,Cost10,Cost11,Cost1,NPV1)
    
    #现金流
    collateral=collateral.rename(columns={'交易编号':'ContractID'})
    Outputs=Collateral_Separate(Outputs,collateral,'前端期权费')
    Outputs=Collateral_Separate(Outputs,collateral,'前端支付')
    Outputs=Collateral_Separate(Outputs,collateral,'部分赎回')
    Outputs=Collateral_Separate(Outputs,collateral,'期间结算')
    Outputs=Collateral_Separate(Outputs,collateral,'定结期间结算')
    Outputs=Collateral_Separate(Outputs,collateral,'其他')
    Outputs=Collateral_Separate(Outputs,collateral,'到期结算')
    Outputs=Collateral_Separate(Outputs,collateral,'全部赎回')
    
    #检查结果
    PnL0=Outputs[Cost0]+Outputs[NPV0]
    PnL1=Outputs[Cost1]+Outputs[NPV1]
    
    Upfront0=Outputs['前端期权费']+Outputs['前端支付']
    Redemption0=Outputs['部分赎回']
    Redemption1=Outputs['全部赎回']
    PayOnExpiry=Outputs['到期结算']
    
    Status1='两期同时存续'
    Status2='上月合约自然到期'
    Status3='上月合约非自然到期'
    Status4='本月新起且存续'
    Status5='本月新起且到期'
    
    flag1=(PnL0.round(decimals=1)==0)
    flag2=(PnL1.round(decimals=1)==0)
    flag3=((Outputs[Cost0]-Outputs[Cost1]-Upfront0-Redemption0).round(decimals=1)==0)
    flag4=((Redemption0+Redemption1+PayOnExpiry).round(decimals=1)==0)
    flag5=((Upfront0+Redemption0+Outputs[Cost1]).round(decimals=1)==0)
    
    Outputs[Status1]=0
    Outputs.loc[(~flag1)&(~flag2)&flag3,[Status1]]=1
    Outputs.loc[(~flag1)&(~flag2)&(~flag3),[Status1]]=100
    
    Outputs[Status2]=0
    Outputs.loc[(Outputs[Cost00]!=0)&flag2,[Status2]]=1
    
    Outputs[Status3]=0
    Outputs.loc[(Outputs[NPV0]!=0)&(Outputs[Status1]+Outputs[Status2]==0)&(~flag4),[Status3]]=1
    Outputs.loc[(Outputs[NPV0]!=0)&(Outputs[Status1]+Outputs[Status2]==0)&flag4,[Status3]]=100
    
    Outputs[Status4]=0
    Outputs.loc[flag1&(~flag2)&flag5,[Status4]]=1
    Outputs.loc[flag1&(~flag2)&(~flag5),[Status4]]=100
    
    Outputs[Status5]=0
    Outputs.loc[flag1&flag2&((Upfront0!=0)|(PayOnExpiry!=0)|(Redemption1!=0)),[Status5]]=1
    Outputs.loc[flag1&flag2&((Upfront0==0)&(PayOnExpiry==0)&(Redemption1==0)),[Status5]]=100
    
    Outputs['Status']='异常'
    Outputs.loc[Outputs[Status1]==1,['Status']]=Status1
    Outputs.loc[Outputs[Status2]==1,['Status']]=Status2
    Outputs.loc[Outputs[Status3]==1,['Status']]=Status3
    Outputs.loc[Outputs[Status4]==1,['Status']]=Status4
    Outputs.loc[Outputs[Status5]==1,['Status']]=Status5
    
    return Outputs


def Check1(lastmonth,thismonth,collateral,FXRate):
    
    thismonth['合约估值']=pd.to_numeric(thismonth['合约估值'],errors='coerce')
    lastmonth['合约估值']=pd.to_numeric(lastmonth['合约估值'],errors='coerce')
    
    
    thismonth['Upfront结算货币']=pd.to_numeric(thismonth['Upfront结算货币'],errors='coerce')
    lastmonth['Upfront结算货币']=pd.to_numeric(lastmonth['Upfront结算货币'],errors='coerce')
    
    thismonth['Upfront结算货币']=thismonth['Upfront结算货币'].replace(np.nan,0.)
    lastmonth['Upfront结算货币']=lastmonth['Upfront结算货币'].replace(np.nan,0.)
    
    FXRate=FXRate[['FROMCCY','FX']]
    thismonth=pd.merge(thismonth,FXRate,how='left',left_on='币种',right_on='FROMCCY')
    lastmonth=pd.merge(lastmonth,FXRate,how='left',left_on='币种',right_on='FROMCCY')
    del thismonth['FROMCCY'],lastmonth['FROMCCY']
    
    thismonth['RMB合约估值']=thismonth['合约估值']*thismonth['FX']
    lastmonth['RMB合约估值']=lastmonth['合约估值']*lastmonth['FX']
    
    
    lastmonth['MATURITYDATEREAL']=pd.to_datetime(lastmonth['MATURITYDATEREAL'])
    thismonth=thismonth.rename(columns={'起始日':'EffectDate'})
    thismonth['EffectDate']=pd.to_datetime(thismonth['EffectDate'])
    
    thismonth=thismonth.rename(columns={'合约编号':'ContractID'})
    lastmonth=lastmonth.rename(columns={'合约编号':'ContractID'})
    
    collateral['确认金额(结算货币)']=pd.to_numeric(collateral['确认金额(结算货币)'],errors='coerce')
    collateral['确认金额(结算货币)']=collateral['确认金额(结算货币)'].replace(np.nan,0.)
    
    return lastmonth,thismonth,collateral
    

def Check0(lastmonth,thismonth,collateral):
    lastmonth_dupl=lastmonth[lastmonth.duplicated(subset='合约编号')]
    thismonth_dupl=thismonth[thismonth.duplicated(subset='合约编号')]
    collateral_dupl=collateral[collateral.duplicated()]
    
    lastmonth=lastmonth.drop_duplicates(subset='合约编号')
    thismonth=thismonth.drop_duplicates(subset='合约编号')
    collateral=collateral.drop_duplicates()
    
    return lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl


if __name__=="__main__":
    path0=os.path.dirname(os.path.realpath(__file__))+'//'
    
    path1='Opt_DM\TheBegin.xlsx'
    path2='Opt_DM\TheEnd.xlsx'
    path3='Opt_DM\Collateral.xlsx'
    path4='Opt_DM\FX.xlsx'
    path5='Opt_DM\Report.xlsx'   
    
    spotdate=pd.to_datetime('2017-11-30')
    lastdate=pd.to_datetime('2017-12-22')
    
    lastmonth=pd.read_excel(path0+path1,0,encoding="gbk",keep_default_na=False)
    thismonth=pd.read_excel(path0+path2,0,encoding="gbk",keep_default_na=False)
    collateral=pd.read_excel(path0+path3,0,encoding="gbk",keep_default_na=False)
    FXRate=pd.read_excel(path0+path4,0,encoding="gbk",keep_default_na=False)
    
    lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl=Check0(lastmonth,thismonth,collateral)
    lastmonth,thismonth,collateral=Check1(lastmonth,thismonth,collateral,FXRate)
    
    Outputs=Check2(lastmonth,thismonth,collateral,lastdate,spotdate)
    
    wbw1=pd.ExcelWriter(path0+path5)
    lastmonth.to_excel(wbw1,'期初表',index=False)
    thismonth.to_excel(wbw1,'期末表',index=False)
    collateral.to_excel(wbw1,'现金流',index=False)
    FXRate.to_excel(wbw1,'FX',index=False)
    Outputs.to_excel(wbw1,'结果',index=False)
    
    if len(lastmonth_dupl)!=0:
        lastmonth_dupl.to_excel(wbw1,'期初表重复',index=False)
    
    if len(thismonth_dupl)!=0:
        thismonth_dupl.to_excel(wbw1,'期末表重复',index=False)
    
    if len(collateral_dupl)!=0:
        collateral_dupl.to_excel(wbw1,'现金流重复',index=False)
    
    wbw1.save()
    