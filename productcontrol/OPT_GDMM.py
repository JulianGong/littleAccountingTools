# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 16:48:50 2017

@author: 028375
"""
import pandas as pd
import numpy as np

spotdate='2017-09-30'
path0='F:\月结表\S201709\股衍\\'.decode('utf-8')


def NoteType(SpotDate,MaturityDateReal,ClientID,Protocols,Books):
    SpotDate=pd.to_datetime(SpotDate)
    MaturityDateReal=pd.to_datetime(MaturityDateReal)
    Result=range(len(ClientID))
    for i in range(len(ClientID)):
        Result[i]=__NoteType(SpotDate,MaturityDateReal[i],ClientID[i],Protocols[i],Books[i])
    Result=pd.DataFrame(Result,columns=['NoteType'])
    return Result


def __NoteType(SpotDate,MaturityDateReal,ClientID,Protocols,Books):
    if MaturityDateReal<=SpotDate:
        Result='已到期'.decode('utf-8')
    elif (ClientID==109966) or (ClientID==109967) or (ClientID==109955):
        Result='非标'.decode('utf-8')
    elif (Protocols==""):
        Result='无合约类型'.decode('utf-8')
    elif ((Protocols=='SAC') or (Protocols=='NAFMII') or (Protocols=='ISDA')) and (Books=='PB'):
        Result='证金跨境期权'.decode('utf-8')
    elif (ClientID==100760) or (ClientID==100062):
        Result='股衍跨境期权'.decode('utf-8')
    elif (Protocols=='SAC') or (Protocols=='NAFMII') or (Protocols=='ISDA'):
        Result='股衍境内期权'.decode('utf-8')
    elif (ClientID==109975) or (ClientID==109980) or (ClientID==109992):
        Result='ELN'
    elif (ClientID==109972) or (ClientID==109901) or (ClientID==109986):
        Result='资金归集'.decode('utf-8')
    elif (ClientID=='109969'):
        Result='资管信托'.decode('utf-8')
    else:
        Result='内部交易'.decode('utf-8')
    
    return Result



def ExpiryType(Books,MaturityDate,MaturityDateReal,KnockOutDate,RempDate):
    Result=range(len(Books))
    for i in range(len(Books)):
        Result[i]=__ExpiryType(Books[Books.index[i]],MaturityDate[MaturityDate.index[i]],MaturityDateReal[MaturityDateReal.index[i]],KnockOutDate[KnockOutDate.index[i]],RempDate[RempDate.index[i]])
    Result=pd.DataFrame(Result,columns=['ExpiryType'], index=Books.index)
    return Result


def __ExpiryType(Books,MaturityDate,MaturityDateReal,KnockOutDate,RempDate):
    if Books=='OTC_FLOW':
        if MaturityDateReal==RempDate:
            return '提前赎回'.decode('utf-8')
        elif MaturityDate==MaturityDateReal:
            return '正常到期'.decode('utf-8')
        else:
            return '异常'.decode('utf-8')
    else:
        if MaturityDateReal==KnockOutDate:
            return '敲出到期'.decode('utf-8')
        elif MaturityDate==MaturityDateReal:
            return '正常到期'.decode('utf-8')
        else:
            return '异常'.decode('utf-8')



def TestTemplate(data):
    path1='全量期权检验模板.xlsx'.decode('utf-8')
    
    data['NoteType']=NoteType(spotdate,data.MATURITYDATEREAL,data['客户编号'.decode('utf-8')],data['合约类型'.decode('utf-8')],data.BOOK)

    tmp0=(data['NoteType']=='已到期'.decode('utf-8'))
    tmp1=(data['产品类型'.decode('utf-8')]=='外部期权'.decode('utf-8'))
    tmp2=(data['NoteType']=='股衍境内期权'.decode('utf-8'))
    tmp3=(data['NoteType']=='股衍跨境期权'.decode('utf-8'))
    
    data0=data[tmp0 & tmp1] #已到期外部期权
    data0=FlowOptionRemp(data0) #增加提前赎回日期
    
    data1=data[tmp2 | tmp3] #外部期权
    data2=data[~((tmp0 & tmp1)|tmp2|tmp3)] #其他合约
    
    [data4,data3]=LongOptionPostDate(data1)
    
    wbw=pd.ExcelWriter(path0+path1)
    data.to_excel(wbw,'全量期权'.decode('utf-8'),index=False)
    
    data0.to_excel(wbw,'已到期外部期权'.decode('utf-8'),index=False)
    data4.to_excel(wbw,'外部期权'.decode('utf-8'),index=False)
    
    data3.to_excel(wbw,'长单期权'.decode('utf-8'),index=False)
    
    wbw.save()
    return data1

    


def ExportOption(data):
    path1='股衍场外期权估值.xlsx'.decode('utf-8')
    path2='股衍ELN资金归集估值表.xlsx'.decode('utf-8')
    data['NoteType']=NoteType(spotdate,data.MATURITYDATEREAL,data['客户编号'.decode('utf-8')],data['合约类型'.decode('utf-8')],data.BOOK)
    
    tmp1=(data['NoteType']=='股衍境内期权'.decode('utf-8'))
    tmp2=(data['NoteType']=='股衍跨境期权'.decode('utf-8'))
    tmp3=(data['NoteType']=='资金归集'.decode('utf-8'))
    data1=data[tmp1] #股衍境内期权
    data2=data[tmp2] #股衍跨境期权
    data3=data[tmp3] #资金归集
    
    del data1['NoteType'], data2['NoteType'], data3['NoteType']
    wbw1=pd.ExcelWriter(path0+path1)
    data1.to_excel(wbw1,'股衍境内期权'.decode('utf-8'),index=False)
    data2.to_excel(wbw1,'股衍跨境期权'.decode('utf-8'),index=False)
    wbw1.save()
    
    wbw2=pd.ExcelWriter(path0+path2)
    data3.to_excel(wbw2,'资金归集'.decode('utf-8'),index=False)
    wbw2.save()


def LongOptionPostDate(data1):
    path1='OptionCashFlow.xlsx'
    PostDate=pd.read_excel(path0+path1,'Sheet1',encoding="gb2312",keep_default_na=False)
    PostDate=PostDate[['ID','CASHFLOWPOSTDATE']]
    PostDate=PostDate.drop_duplicates()
    #print PostDate
    
    data2=data1[data1['合约估值'.decode('utf-8')]=='']
    data1=data1[data1['合约估值'.decode('utf-8')]<>'']
    data2=pd.merge(data2,PostDate,how='left',left_on='合约编号'.decode('utf-8'),right_on='ID')
    return data1, data2


def FlowOptionRemp(data):
    path1='FlowTradeDetail.xlsx'
    
    ftd=pd.read_excel(path0+path1,'Sheet1',encoding="gb2312",keep_default_na=False)
    ftd=ftd[['TRADEID','NOTIONAL','TRADEDATE']]
    ftd=ftd[ftd.NOTIONAL==0]
    data=pd.merge(data,ftd,how='left',left_on='合约编号'.decode('utf-8'),right_on='TRADEID')
    del data['TRADEID'], data['NOTIONAL']
    data=data.rename(columns={'TRADEDATE': 'FlowOptRempDate'})
    
    data['ExpiryType']=ExpiryType(data.BOOK,data.MATURITYDATE,data.MATURITYDATEREAL,data['敲出日期'.decode('utf-8')],data.FlowOptRempDate)
    return data



if __name__ == "__main__":
    path1='全量期权估值表.xlsx'.decode('utf-8')
    data=pd.read_excel(path0+path1,'Sheet1',encoding="gb2312",keep_default_na=False)
    
    tmp1=data[data['合约编号'.decode('utf-8')]=='SACTC0419A20170120_put']
    tmp2=data[data['合约编号'.decode('utf-8')]=='SACTC0419A20170120_mountain_range_call']
    tmp2['合约估值'.decode('utf-8')]=np.array(tmp1['合约估值'.decode('utf-8')])+np.array(tmp2['合约估值'.decode('utf-8')])
    
    data=data.drop([tmp1.index[0],tmp2.index[0]])
    data=data.append(tmp2)
    
    data['index1']=range(len(data))
    data=data.set_index(['index1'])
    
    data1=TestTemplate(data)
    ExportOption(data)