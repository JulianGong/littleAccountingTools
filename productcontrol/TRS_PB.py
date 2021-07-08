# -*- coding: utf-8 -*-
"""
Created on Wed Jul 12 17:15:10 2017

@author: 028375
"""


#import pandas as pd
from pandas import merge,ExcelWriter,DataFrame,read_excel,read_csv
from numpy import sum

spotdate='20171018'
lastdate='20171017'
a=1.
path0="F:\Product Control\月结表\境内TRS\S201710\\".decode('utf-8')



def TestTemplate1(Status,Collateral,Position,LastStatus):
    LastStatus=LastStatus[['ACCOUNTID','COLLATERALAMT_CASH','CS_CONTRACT_VALUE']]
    LastStatus.columns=[['TradeID','LastCollateral','LastValue']]
    Status=merge(Status,LastStatus,how='outer',left_on='ACCOUNTID',right_on='TradeID')
    
    Result=range(len(Status))
    for i in range(len(Status)):
        tmp1=Position[Position['ACCOUNTID']==Status['ACCOUNTID'][Status.index[i]]]['CURRENTQTY']
        tmp2=Position[Position['ACCOUNTID']==Status['ACCOUNTID'][Status.index[i]]]['MTMPRICE']
        Result[i]=sum(tmp1*tmp2)
        
    Result=DataFrame(Result,columns=['Position'],index=Status.index)
    Status['Position']=Result['Position']
    return Status



def TestTemplate2(Status,Collateral,Position,PTH,Info):
    path1=('证金境内TRS检验'+spotdate+'.xlsx').decode('utf-8')
    path2=('证金境内TRS估值'+lastdate+'.xlsx').decode('utf-8')
    #path1=('证金境内TRS检验'+spotdate+'_new1'+'.xlsx').decode('utf-8')
    #path2=('证金境内TRS估值'+lastdate+'_new1'+'.xlsx').decode('utf-8')
    
    LastStatus=read_excel(path0+path2,'Status',encoding="gb2312",keep_default_na=False)
    LastStatus=LastStatus[LastStatus['ACCOUNTID']<>'']
    
    Status1=TestTemplate1(Status,Collateral,Position,LastStatus)
    
    LastCollateral=read_excel(path0+path2,'Collateral',encoding="gb2312",keep_default_na=False)
    Collateral=Collateral.append(LastCollateral)
    
    Test=LastStatus[['ACCOUNTID','FINANCINGMETHOD_OS','RESETINTERVAL','COLLATERALAMT_SCC','UNPAID_FINANCING','REALIZEDPNL_SCC','COLLATERAL_REBATE','PTHFEE']]
    
    IntRate=Info[['ACCOUNTID', ' "GNLRESET"',' "SPREAD"',' "COLLREBATESPREAD"']]
    Test=merge(Test,IntRate,how='left',left_on='ACCOUNTID',right_on='ACCOUNTID')
    
    tmp0=Status[['ACCOUNTID','COLLATERALAMT_SCC']]
    tmp0=tmp0.rename(columns={'COLLATERALAMT_SCC' : 'CollateralAMT2'})
    Test=merge(Test,tmp0,how='left',left_on='ACCOUNTID',right_on='ACCOUNTID')
    
    Notional=range(len(Test))
    for i in range(len(Test)):
        tmp1=Position[Position['ACCOUNTID']==Test['ACCOUNTID'][Test.index[i]]]['NOTIONAL']
        Notional[i]=sum(tmp1)
        
    Notional=DataFrame(Notional,columns=['Notional'],index=Test.index)
    Test['Notional']=Notional['Notional']
    
    
    DailyFinFee=range(len(Test)); DailyRebate=range(len(Test))
    for i in range(len(Test)):
        tmp0=Test['FINANCINGMETHOD_OS'][Test.index[i]]
        if tmp0<>'CONTRACT':
            tmp1=Test['Notional'][Test.index[i]]
            
            tmp2=Test['CollateralAMT2'][Test.index[i]]
            if tmp2=='': tmp2=0.
            
            tmp3=Test[' "SPREAD"'][Test.index[i]]
            if tmp3=='': tmp3=0.
            
            DailyFinFee[i]=a*max(tmp1-tmp2,0.)*tmp3/360. ;DailyRebate[i]=0.
        else:
            tmp1=Test['Notional'][Test.index[i]]
            tmp2=Test['CollateralAMT2'][Test.index[i]]
            if tmp2=='': tmp2=0.
            
            tmp3=Test[' "SPREAD"'][Test.index[i]]
            if tmp3=='': tmp3=0.
            
            tmp4=Test[' "COLLREBATESPREAD"'][Test.index[i]]
            if tmp4=='': tmp4=0.
            
            DailyFinFee[i]=a*tmp1*tmp3/360. ; DailyRebate[i]=a*tmp2*float(tmp4)/360.
            
    DailyFinFee=DataFrame(DailyFinFee,columns=['DailyFinFee'],index=Test.index)
    DailyRebate=DataFrame(DailyRebate,columns=['DailyRebate'],index=Test.index)
    
    Test['DailyFinFee']=DailyFinFee['DailyFinFee']; Test['DailyRebate']=DailyRebate['DailyRebate']
    
     
    DailyPTHFee=range(len(Test))
    for i in range(len(Test)):
        tmp1=PTH[PTH['ACCOUNTID']==Test['ACCOUNTID'][Test.index[i]]]['BORROWQTY']
        tmp2=PTH[PTH['ACCOUNTID']==Test['ACCOUNTID'][Test.index[i]]]['MTMPRICE']
        tmp3=PTH[PTH['ACCOUNTID']==Test['ACCOUNTID'][Test.index[i]]]['SLRATE']
        DailyPTHFee[i]=a*sum(tmp1*tmp2*tmp3/360.)
    DailyPTHFee=DataFrame(DailyPTHFee,columns=['DailyPTHFee'],index=Test.index)
    Test['DailyPTHFee']=DailyPTHFee['DailyPTHFee']
    
    tmp1=Status[['ACCOUNTID','REALIZEDPNL_SCC']]
    tmp1=tmp1.rename(columns={'REALIZEDPNL_SCC' : 'RealizedPnL_SCC2'})
    Test=merge(Test,tmp1,how='left',left_on='ACCOUNTID',right_on='ACCOUNTID')
    
    Test['ResetAmount']=LastStatus['RESET_AMOUNT']
    
    UnRealizedPnL=range(len(Test))
    for i in range(len(Test)):
        tmp1=Position[Position['ACCOUNTID']==Test['ACCOUNTID'][Test.index[i]]]['UNREALIZEDPNL']
        UnRealizedPnL[i]=sum(tmp1)
    UnRealizedPnL=DataFrame(UnRealizedPnL,columns=['UnRealizedPnL'],index=Test.index)
    Test['UnRealizedPnL']=UnRealizedPnL['UnRealizedPnL']
    
    Test['ContractValue']=-(Test['UNPAID_FINANCING']-Test['DailyFinFee'])-(Test['PTHFEE']-Test['DailyPTHFee'])-(Test['RealizedPnL_SCC2']+Test['UnRealizedPnL'])
    Test['ContractValue']=Test['ContractValue']-(Test['COLLATERAL_REBATE']+Test['DailyRebate'])
    
    ContractValue=Status[['ACCOUNTID','CS_CONTRACT_VALUE']]
    Test=merge(Test,ContractValue,how='outer',left_on='ACCOUNTID',right_on='ACCOUNTID')
    
    
    wbw=ExcelWriter(path0+path1)
    
    Test.to_excel(wbw,'Test',index=False)
    Status1.to_excel(wbw,'Status1',index=False)
    Status.to_excel(wbw,'Status',index=False)
    Collateral.to_excel(wbw,'Collateral',index=False)
    Position.to_excel(wbw,'Position',index=False)
    PTH.to_excel(wbw,'PTH',index=False)
    Info.to_excel(wbw,'Info',index=False)
    
    wbw.save()
    


def ExportTRS(Status,Collateral):
    path1=('证金境内TRS估值'+spotdate+'.xlsx').decode('utf-8')
    path2=('证金境内TRS估值'+lastdate+'.xlsx').decode('utf-8')
    #path1=('证金境内TRS估值'+spotdate+'_new1'+'.xlsx').decode('utf-8')
    #path2=('证金境内TRS估值'+lastdate+'_new1'+'.xlsx').decode('utf-8')
    
    LastCollateral=read_excel(path0+path2,'Collateral',encoding="gb2312",keep_default_na=False)
    Collateral=Collateral.append(LastCollateral)
    
    wbw=ExcelWriter(path0+path1)
    Status.to_excel(wbw,'Status'.decode('utf-8'),index=False)
    Collateral.to_excel(wbw,'Collateral'.decode('utf-8'),index=False)
    wbw.save()
    
    
    

if __name__ == "__main__":
    path1='FIN_ACCOUNTSTATUS_OS_'+spotdate+'.csv'
    path2='FIN_COLLATERALFLOW_OS_'+spotdate+'.csv'
    path3='FIN_POSITIONDETAIL_OS_'+spotdate+'.csv'
    #path1='accountstatus_'+spotdate+'.csv'
    #path2='CollateralFlow_'+spotdate+'.csv'
    #path3='PositionDetail_'+spotdate+'.csv'
    
    path4='FIN_PTHPOSITIONDETAIL_OS_'+spotdate+'.csv'
    path5='FIN_RESETDATES_OS_'+spotdate+'.csv'
    path6='FIN_CLIENTINFO_OS_'+spotdate+'.csv'    
    
    Status=read_csv(path0+path1,encoding="gb2312",keep_default_na=False)
    Collateral=read_csv(path0+path2,encoding="gb2312",keep_default_na=False)
    Position=read_csv(path0+path3,encoding="gb2312",keep_default_na=False)
    PTH=read_csv(path0+path4,encoding="gb2312",keep_default_na=False)
    #ResetDate=pd.read_csv(path0+path5,encoding="gb2312",keep_default_na=False)
    Info=read_csv(path0+path6,keep_default_na=False)
    del Info[' "ACCOUNTNAME"']
    
    Status=Status[Status['CLIENT_TYPE']<>'CS测试账户'.decode('utf-8')]
  
    ExportTRS(Status,Collateral)
    TestTemplate2(Status,Collateral,Position,PTH,Info)
    