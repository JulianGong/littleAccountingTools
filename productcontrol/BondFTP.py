# -*- coding: utf-8 -*-
"""
Created on Thu Jul 13 15:44:53 2017

@author: 028375
"""

from __future__ import unicode_literals, division,print_function

import pandas as pd
import numpy as np

path0='F:\CIT_FC_BB_REPLY_20171016092016.csv'

tradedate=['2017-10-6','2017-10-9','2017-10-10','2017-10-11','2017-10-12','2017-10-13']
tradedate=pd.to_datetime(tradedate)

BondPrice=pd.read_csv(path0,keep_default_na=False)

BondISIN=BondPrice[BondPrice.columns[0]].unique()

BondPrice=BondPrice[BondPrice['TRADE_DATE']<>' '] #故意留空
BondPrice['TRADE_DATE']=pd.to_datetime(BondPrice['TRADE_DATE'])


BGNPrice=np.zeros([len(BondISIN),len(tradedate)])
CTCSPrice=np.zeros([len(BondISIN),len(tradedate)])
OutPrice=np.zeros([len(BondISIN),len(tradedate)])

for j in range(len(tradedate)):
    for i in range(len(BondISIN)):
        tmp1=(BondPrice[BondPrice.columns[0]]==BondISIN[i])
        tmp2=(BondPrice[BondPrice.columns[2]]==tradedate[j])
        tmp3=(BondPrice[BondPrice.columns[1]]=='CTCS')
        tmp4=(BondPrice[BondPrice.columns[1]]<>'CTCS')
        
        CTCSPrice[i,j]=np.sum(np.array(BondPrice[tmp1&tmp2&tmp3]['PX_LAST']))
        BGNPrice[i,j]=np.sum(np.array(BondPrice[tmp1&tmp2&tmp4]['PX_LAST']))
        if CTCSPrice[i,j]<>0:
            OutPrice[i,j]=CTCSPrice[i,j]
        else:
            OutPrice[i,j]=BGNPrice[i,j]

CTCSPrice=pd.DataFrame(CTCSPrice,columns=tradedate,index=BondISIN)
BGNPrice=pd.DataFrame(BGNPrice,columns=tradedate,index=BondISIN)
OutPrice=pd.DataFrame(OutPrice,columns=tradedate,index=BondISIN).transpose()

path1=('F:\境内股票债券估值.xlsx')
wbw=pd.ExcelWriter(path1)
OutPrice.to_excel(wbw,'Sheet1',index=True)
wbw.save()
