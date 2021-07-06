# -*- coding: utf-8 -*-
"""
Created on Thu Feb 22 11:05:21 2018

@author: 028375
"""

from __future__ import unicode_literals, division
import pandas as pd
import os.path
import numpy as np


def Check2(lastmonth,thismonth,collateral):
    ContractID=(thismonth['ContractID'].append(lastmonth['ContractID'])).append(collateral['ContractID']).drop_duplicates()
    Outputs=pd.DataFrame(ContractID).reset_index(drop=True)
    
    cost0=lastmonth[['ContractID','期权标的','标的类型','Upfront结算货币']]
    Outputs=pd.merge(Outputs,cost0,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Upfront结算货币':'期初表Upfront','期权标的':'期初表期权标的','标的类型':'期初表标的类型'})
    
    cost1=thismonth[['ContractID','期权标的','标的类型','Upfront结算货币']]
    Outputs=pd.merge(Outputs,cost1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Upfront结算货币':'期末表Upfront','期权标的':'期末表期权标的','标的类型':'期末表标的类型'})
    
    tmp1=collateral.groupby(['ContractID'])[['期权标的','标的类型']].first().reset_index()
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'期权标的':'资金表期权标的','标的类型':'资金表标的类型'})
    
    collateral1=collateral.groupby(['ContractID','现金流类型'])['确认金额(结算货币)'].sum().reset_index()
    collateral1=collateral1.rename(columns={'现金流类型':'CashType','确认金额(结算货币)':'Amount'})
    
    tmp1=collateral1[collateral1['CashType']=='前端支付'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'前端支付'})
    
    tmp1=collateral1[collateral1['CashType']=='前端期权费'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'前端期权费'})
    
    tmp1=collateral1[collateral1['CashType']=='展期期权费'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'展期期权费'})
    
    tmp1=collateral1[collateral1['CashType']=='到期结算'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'到期结算'})
    
    tmp1=collateral1[collateral1['CashType']=='部分赎回'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'部分赎回'})
    
    tmp1=collateral1[collateral1['CashType']=='全部赎回'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'全部赎回'})
    
    tmp1=collateral1[collateral1['CashType']=='期间结算'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'期间结算'})
    
    tmp1=collateral1[collateral1['CashType']=='红利支付'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'红利支付'})
    
    tmp1=collateral1[collateral1['CashType']=='其他'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'其他'})
    
    tmp1=collateral1[collateral1['CashType']=='定结期间结算'][['ContractID','Amount']]
    Outputs=pd.merge(Outputs,tmp1,how='left',on='ContractID')
    Outputs=Outputs.rename(columns={'Amount':'定结期间结算'})
    
    Outputs['status1']=''
    flag1=np.isnan(Outputs['期初表Upfront'])
    flag2=np.isnan(Outputs['期末表Upfront'])
    Outputs.loc[flag1&flag2,['status1']]='新起到期'
    Outputs.loc[(~flag1)&flag2,['status1']]='存续到期'
    Outputs.loc[flag1&(~flag2),['status1']]='新起存续'
    Outputs.loc[(~flag1)&(~flag2),['status1']]='两期存续'
    
    Outputs['status2']=''
    flag1=(Outputs['status1']=='新起到期')
    flag2=(Outputs['status1']=='存续到期')
    flag3=(Outputs['status1']=='新起存续')
    flag4=(Outputs['status1']=='两期存续')
    
    colflag1=np.isnan(Outputs['前端支付'])
    colflag2=np.isnan(Outputs['前端期权费'])
    colflag3=np.isnan(Outputs['展期期权费'])
    colflag4=np.isnan(Outputs['到期结算'])
    colflag5=np.isnan(Outputs['全部赎回'])
    colflag6=np.isnan(Outputs['部分赎回'])
    colflag7=np.isnan(Outputs['定结期间结算']) #update 0.2.3
    
    tmp1=Outputs[['ContractID','期初表Upfront','期末表Upfront','前端支付','前端期权费','展期期权费','到期结算','部分赎回','全部赎回','定结期间结算']]
    tmp1=tmp1.replace(np.nan,0.)
    flag5=(tmp1['期末表Upfront']!=0)
    flag6=(tmp1['期末表Upfront']-tmp1['期初表Upfront']).round(decimals=4)==0
    flag7=(tmp1['期末表Upfront']-tmp1['前端支付']).round(decimals=4)==0
    flag8=(tmp1['期末表Upfront']-(tmp1['前端期权费']+tmp1['展期期权费']+tmp1['部分赎回'])).round(decimals=4)==0
    #flag9=(tmp1['期末表Upfront']-(tmp1['期初表Upfront']+tmp1['展期期权费']+tmp1['部分赎回'])).round(decimals=4)==0 #update 0.2.3
    flag9=(tmp1['期末表Upfront']-(tmp1['期初表Upfront']+tmp1['展期期权费']+tmp1['部分赎回']+tmp1['定结期间结算'])).round(decimals=4)==0 # update 0.2.3 增加定结期间结算 
    
    #新起到期
    Outputs.loc[flag1,['status2']]='流水异常'
    # Outputs.loc[flag1&((~colflag1)|(~colflag2))&((~colflag4)|(~colflag5)),['status2']]='流水正常' #update 0.2.3
    Outputs.loc[flag1&((~colflag4)|(~colflag5)),['status2']]='流水正常'  #update 0.2.3
    
    #存续到期
    Outputs.loc[flag2,['status2']]='流水异常'
    Outputs.loc[flag2&((~colflag4)|(~colflag5)),['status2']]='流水正常'
    
    #新起存续
    Outputs.loc[flag3,['status2']]='流水异常'
    Outputs.loc[flag3&flag5&((~colflag1)|(~colflag2))&colflag4&colflag5,['status2']]='流水正常'
    tmp_flag=((~colflag1)&tmp1['前端支付']!=0)|((~colflag2)&tmp1['前端期权费']!=0) #前端支付/前端期权费存在,且不等于0
    Outputs.loc[flag3&(~flag5)&(colflag4&colflag5)&(~tmp_flag),['status2']]='流水正常'
    
    #两期存续
    Outputs.loc[flag4,['status2']]='流水异常'
    Outputs.loc[flag4&flag6&(colflag3&colflag6&colflag4&colflag5),['status2']]='流水正常'
    # Outputs.loc[flag4&(~flag6)&((~colflag3)|(~colflag6)&colflag4&colflag5),['status2']]='流水正常' #update 0.2.3
    Outputs.loc[flag4&(~flag6)&((~colflag3)|(~colflag6)|(~colflag7)&colflag4&colflag5),['status2']]='流水正常' #增加定结期间结算 #update 0.2.3
    
    Outputs['status3']=''
    flag10=(Outputs['status2']=='流水异常')
    Outputs.loc[flag10,['status3']]='流水异常,未验证金额'
    Outputs.loc[(~flag10)&flag1,['status3']]='无需验证金额'
    Outputs.loc[(~flag10)&flag2,['status3']]='无需验证金额'
    
    Outputs.loc[(~flag10)&flag3,['status3']]='金额异常'
    Outputs.loc[(~flag10)&flag3&(flag7|flag8|(~flag5)),['status3']]='金额正常'
    
    Outputs.loc[(~flag10)&flag4,['status3']]='金额异常'
    Outputs.loc[(~flag10)&flag4&(flag6|flag9),['status3']]='金额正常'
    return Outputs
    

def Check1(lastmonth,thismonth,collateral):
    
    thismonth['Upfront结算货币']=pd.to_numeric(thismonth['Upfront结算货币'],errors='coerce')
    lastmonth['Upfront结算货币']=pd.to_numeric(lastmonth['Upfront结算货币'],errors='coerce')
    thismonth['Upfront结算货币']=thismonth['Upfront结算货币'].replace(np.nan,0.)
    lastmonth['Upfront结算货币']=lastmonth['Upfront结算货币'].replace(np.nan,0.)
    
    lastmonth['MATURITYDATEREAL']=pd.to_datetime(lastmonth['MATURITYDATEREAL'])
    thismonth=thismonth.rename(columns={'起始日':'EffectDate'})
    thismonth['EffectDate']=pd.to_datetime(thismonth['EffectDate'])
    
    thismonth=thismonth.rename(columns={'合约编号':'ContractID'})
    lastmonth=lastmonth.rename(columns={'合约编号':'ContractID'})
    collateral=collateral.rename(columns={'交易编号':'ContractID'})
    
    collateral['现金流产生日期']=pd.to_datetime(collateral['现金流产生日期'])
    collateral['确认金额(结算货币)']=pd.to_numeric(collateral['确认金额(结算货币)'],errors='coerce')
    collateral['确认金额(结算货币)']=collateral['确认金额(结算货币)'].replace(np.nan,0.)
    
    return lastmonth,thismonth,collateral



def Check0(lastmonth,thismonth,collateral):
    lastmonth_dupl=lastmonth[lastmonth.duplicated(subset='合约编号')]
    thismonth_dupl=thismonth[thismonth.duplicated(subset='合约编号')]
    collateral_dupl=collateral[collateral.duplicated()]
    
    lastmonth=lastmonth.drop_duplicates(subset='合约编号')
    thismonth=thismonth.drop_duplicates(subset='合约编号')
    collateral=collateral.drop_duplicates(subset=['交易编号','现金流类型','现金流产生日期','确认金额(结算货币)'])
    
    flag1=collateral['现金流类型']!='前端支付'
    flag2=collateral['现金流类型']!='前端期权费'
    flag3=collateral['现金流类型']!='展期期权费'
    flag4=collateral['现金流类型']!='到期结算'
    flag5=collateral['现金流类型']!='部分赎回'
    flag6=collateral['现金流类型']!='全部赎回'
    flag7=collateral['现金流类型']!='期间结算'
    flag8=collateral['现金流类型']!='红利支付'
    flag9=collateral['现金流类型']!='其他'
    
    flag10=collateral['现金流类型']!='定结期间结算'
    collateral_newtype=collateral[flag1&flag2&flag3&flag4&flag5&flag6&flag7&flag8&flag9&flag10]
    return lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl,collateral_newtype



if __name__=="__main__":
    path0=os.path.dirname(os.path.realpath(__file__))+'//'
    spotdate=pd.to_datetime('2017-11-30')
    lastdate=pd.to_datetime('2017-12-22')
    
    path1='Opt_DM\TheBegin.xlsx'
    path2='Opt_DM\TheEnd.xlsx'
    path3='Opt_DM\Collateral.xlsx'
    path4='Opt_DM\Report3.xlsx'   
    
    lastmonth=pd.read_excel(path0+path1,0,encoding="gbk",keep_default_na=False)
    thismonth=pd.read_excel(path0+path2,0,encoding="gbk",keep_default_na=False)
    collateral=pd.read_excel(path0+path3,0,encoding="gbk",keep_default_na=False)
    
    lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl,collateral_newtype=Check0(lastmonth,thismonth,collateral)
    lastmonth,thismonth,collateral=Check1(lastmonth,thismonth,collateral)
    
    Outputs=Check2(lastmonth,thismonth,collateral)
    
    wbw1=pd.ExcelWriter(path0+path4)
    lastmonth.to_excel(wbw1,'期初表',index=False)
    thismonth.to_excel(wbw1,'期末表',index=False)
    collateral.to_excel(wbw1,'现金流',index=False)
    Outputs.to_excel(wbw1,'结果',index=False)
    
    if len(lastmonth_dupl)!=0:
        lastmonth_dupl.to_excel(wbw1,'期初表重复',index=False)
    
    if len(thismonth_dupl)!=0:
        thismonth_dupl.to_excel(wbw1,'期末表重复',index=False)
    
    if len(collateral_dupl)!=0:
        collateral_dupl.to_excel(wbw1,'现金流重复',index=False)
        
    if len(collateral_newtype)!=0:
        collateral_newtype.to_excel(wbw1,'新现金流类型',index=False)
    
    wbw1.save()
    
    