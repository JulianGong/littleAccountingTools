# -*- coding: utf-8 -*-
"""
Created on Wed Jul 12 11:00:56 2017

@author: 028375
"""

import pandas as pd
import numpy as np

begindate='20171001'
spotdate='20171018'
lastdate='20171017'
path0='F:\月结表\境内TRS\S201710\\'.decode('utf-8')


def TestTemplate(Status,Collateral,Position):
    path1=('股衍境内TRS检验'+spotdate+'.xlsx').decode('utf-8')
    path2=('股衍境内TRS估值'+lastdate+'.xlsx').decode('utf-8')
    
    LastStatus=pd.read_excel(path0+path2,'账户状态'.decode('utf-8'),encoding="gb2312",keep_default_na=False)
    LastStatus=LastStatus[['清单编号'.decode('utf-8'),'客户总预付金'.decode('utf-8'),'我方角度合约价值'.decode('utf-8')]]
    LastStatus.columns=[['TradeID','LastCollateral','LastValue']]
    
    Status=pd.merge(Status,LastStatus,how='outer',left_on='清单编号'.decode('utf-8'),right_on='TradeID')

    Result=range(len(Status))
    for i in range(len(Status)):
        tmp1=Position[Position['合约编号'.decode('utf-8')]==Status['清单编号'.decode('utf-8')][Status.index[i]]]['持仓数量'.decode('utf-8')]
        tmp2=Position[Position['合约编号'.decode('utf-8')]==Status['清单编号'.decode('utf-8')][Status.index[i]]]['最新价'.decode('utf-8')]
        Result[i]=np.sum(tmp1*tmp2)
        
    Result=pd.DataFrame(Result,columns=['Position'],index=Status.index)
    Status['Position']=Result['Position']
    
    wbw=pd.ExcelWriter(path0+path1)
    Status.to_excel(wbw,'Status',index=False)
    Collateral.to_excel(wbw,'Collateral',index=False)
    Position.to_excel(wbw,'Position',index=False)
    wbw.save()
    
    return Status



def ExportTRS(Status,Collateral):
    path1=('股衍境内TRS估值'+spotdate+'.xlsx').decode('utf-8')
    
    wbw=pd.ExcelWriter(path0+path1)
    Status.to_excel(wbw,'账户状态'.decode('utf-8'),index=False)
    Collateral.to_excel(wbw,'资金流水'.decode('utf-8'),index=False)
    wbw.save()
    
    

if __name__ == "__main__":
    path1='收益互换日终报表-账户状态.xlsx'.decode('utf-8')
    path2='收益互换日终报表-资金流水表.xlsx'.decode('utf-8')
    path3='收益互换日终报表-组合持仓.xlsx'.decode('utf-8')
    
    Status=pd.read_excel(path0+path1,'Sheet1',encoding="gb2312",keep_default_na=False)
    Collateral=pd.read_excel(path0+path2,'Sheet1',encoding="gb2312",keep_default_na=False)
    Position=pd.read_excel(path0+path3,'Sheet1',encoding="gb2312",keep_default_na=False)
    
    Status=Status[Status['交易状态'.decode('utf-8')]==1]
    Collateral=Collateral[pd.to_datetime(Collateral['日期'.decode('utf-8')])>=pd.to_datetime(begindate)]
    Position=Position[Position['合约状态'.decode('utf-8')]==1]
    
    ExportTRS(Status,Collateral)
    Status0=TestTemplate(Status,Collateral,Position)
    