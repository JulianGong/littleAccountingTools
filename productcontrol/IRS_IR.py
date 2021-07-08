# -*- coding: utf-8 -*-
"""
Created on Wed Jul 12 11:00:56 2017

@author: 028375
"""

import sys
sys.path.append('f:\mycode\Python\Boost') 
import pandas as pd
import IRS
import matlab.engine
import iDate
import numpy as np

SpotDate='21-Aug-2017'
SpotDate=[iDate.__date2num(iDate.__str2date(SpotDate))+1]
path0='F:\月结表\S201708\\'.decode('utf-8')


class xIR_IRSPricer_Class(IRS.IRSPricer_Class2):
    
    def __init__(self,Index_Shibor3M,Index_FR007,Index_FDR007,Curve_Shibor3M,Curve_FR007,Curve_FDR007,eng):
        Index_Shibor3M=(np.array(Index_Shibor3M)).tolist()
        Index_FR007=(np.array(Index_FR007)).tolist()
        Index_FDR007=(np.array(Index_FDR007)).tolist()
        Curve_Shibor3M=(np.array(Curve_Shibor3M)).tolist()
        Curve_FR007=(np.array(Curve_FR007)).tolist()
        Curve_FDR007=(np.array(Curve_FDR007)).tolist()
        
        self.Index_Shibor3M=Index_Shibor3M
        self.Index_FR007=Index_FR007
        self.Index_FDR007=Index_FDR007
        self.Curve_Shibor3M=Curve_Shibor3M
        self.Curve_FR007=Curve_FR007
        self.Curve_FDR007=Curve_FDR007
        self.eng=eng
       
    def Pricing2(self,BuyFlag,Notional,EffectDate,ExpiryDate,SpotDate,Strike,RefRate,Spread):
        eng=self.eng
        
        if RefRate=='SHIBOR-3M':
            RefCurve=self.Curve_Shibor3M
            DisCurve=self.Curve_FR007
            RefRate_his=self.Index_Shibor3M
            NPV=self.Shibor3M(BuyFlag,Notional,EffectDate,ExpiryDate,SpotDate,Strike,Spread,RefCurve,DisCurve,RefRate_his,eng)
        elif RefRate=='FR007':
            RefCurve=self.Curve_FR007
            DisCurve=self.Curve_FR007
            RefRate_his=self.Index_FR007
            NPV=self.FR007(BuyFlag,Notional,EffectDate,ExpiryDate,SpotDate,Strike,Spread,RefCurve,DisCurve,RefRate_his,eng)
        elif RefRate=='FDR007':
            RefCurve=self.Curve_FDR007
            DisCurve=self.Curve_FR007
            RefRate_his=self.Index_FDR007
            NPV=self.FDR007(BuyFlag,Notional,EffectDate,ExpiryDate,SpotDate,Strike,Spread,RefCurve,DisCurve,RefRate_his,eng)
        return NPV
        


if __name__ == "__main__":
    path1='固收利率互换估值检验.xlsx'.decode('utf-8')
    
    Contract_IRS=pd.read_excel(path0+path1,'IRS',encoding="gb2312",keep_default_na=False)
    Index_Shibor3M=pd.read_excel(path0+path1,'Index_Shibor3M',encoding="gb2312",keep_default_na=False)
    Index_FR007=pd.read_excel(path0+path1,'Index_FR007',encoding="gb2312",keep_default_na=False)
    Index_FDR007=pd.read_excel(path0+path1,'Index_FDR007',encoding="gb2312",keep_default_na=False)
    Curve_Shibor3M=pd.read_excel(path0+path1,'Curve_Shibor3M',encoding="gb2312",keep_default_na=False)
    Curve_FR007=pd.read_excel(path0+path1,'Curve_FR007',encoding="gb2312",keep_default_na=False)
    Curve_FDR007=pd.read_excel(path0+path1,'Curve_FDR007',encoding="gb2312",keep_default_na=False)
    
    Contract_IRS=Contract_IRS[['交易编号'.decode('utf-8'),'名义本金（万元）'.decode('utf-8'),'交易方向'.decode('utf-8'),'支付利率'.decode('utf-8'),'收取利率'.decode('utf-8'),'起息日期'.decode('utf-8'),'到期日期'.decode('utf-8')]]
    Index_Shibor3M=Index_Shibor3M[['日期'.decode('utf-8'),'利率(%)'.decode('utf-8')]]
    Index_FR007=Index_FR007[['日期'.decode('utf-8'),'利率(%)'.decode('utf-8')]]
    Index_FDR007=Index_FDR007[['日期'.decode('utf-8'),'利率(%)'.decode('utf-8')]]
    Curve_Shibor3M=Curve_Shibor3M[['关键点'.decode('utf-8'),'收益率(%)'.decode('utf-8')]]
    Curve_FR007=Curve_FR007[['关键点'.decode('utf-8'),'收益率(%)'.decode('utf-8')]]
    Curve_FDR007=Curve_FDR007[['关键点'.decode('utf-8'),'收益率(%)'.decode('utf-8')]]
    
    tmp1=Contract_IRS['交易方向'.decode('utf-8')]=='支付固定利率'.decode('utf-8')
    tmp2=Contract_IRS['交易方向'.decode('utf-8')]=='支付浮动利率'.decode('utf-8')
    Contract_IRS['Notional']=Contract_IRS['名义本金（万元）'.decode('utf-8')]*10000.
    
    Contract_IRS['BuyFlag']=0
    Contract_IRS['BuyFlag'][tmp1]=-1
    Contract_IRS['BuyFlag'][tmp2]=1
    
    Contract_IRS['FixRate']=0
    Contract_IRS['FixRate'][tmp1]=Contract_IRS['支付利率'.decode('utf-8')][tmp1]
    Contract_IRS['FixRate'][tmp2]=Contract_IRS['收取利率'.decode('utf-8')][tmp2]
    Contract_IRS['FixRate']=Contract_IRS['FixRate']/100.
    
    Contract_IRS['FltRate']=0
    Contract_IRS['FltRate'][tmp1]=Contract_IRS['收取利率'.decode('utf-8')][tmp1]
    Contract_IRS['FltRate'][tmp2]=Contract_IRS['支付利率'.decode('utf-8')][tmp2]
    
    eng=matlab.engine.start_matlab()
    xIR_IRSPricer=xIR_IRSPricer_Class(Index_Shibor3M,Index_FR007,Index_FDR007,Curve_Shibor3M,Curve_FR007,Curve_FDR007,eng)
    
    N=len(Contract_IRS['FixRate'])
    N=10
    Result=range(N)
    for i in range(N):
        BuyFlag=[Contract_IRS['BuyFlag'][i]]
        Notional=[Contract_IRS['Notional'][i]]
        EffectDate=[Contract_IRS['起息日期'.decode('utf-8')][i]]
        ExpiryDate=[Contract_IRS['到期日期'.decode('utf-8')][i]]
        Strike=[Contract_IRS['FixRate'][i]]
        RefRate=Contract_IRS['FltRate'][i]
        Spread=[0.]
        Result[i]=xIR_IRSPricer.Pricing2(BuyFlag,Notional,EffectDate,ExpiryDate,SpotDate,Strike,RefRate,Spread)
        
    eng.quit
    
    