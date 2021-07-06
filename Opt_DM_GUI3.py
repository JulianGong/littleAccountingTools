# -*- coding: utf-8 -*-
"""
Created on Thu Feb 22 09:52:09 2018

@author: 028375
"""

from __future__ import unicode_literals, division
try:
    import tkinter as tk
    import tkinter.messagebox as msbox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as msbox
    
import Opt_DM_Script3 as GDMM
import os
import pandas as pd
import numpy as np

try:
    #path0=os.path.dirname(os.path.realpath(__file__))+'\\' 
    path0=os.getcwd()+'\\' 
except:
    #path0=os.path.dirname((os.path.realpath(__file__)).decode('gb2312'))+'\\'
    path0=(os.getcwd()).decode('gb2312')+'\\' 
    


#def check_contain_chinese(check_str):
#    for ch in check_str:
#        if ord(ch)>127:
#            return True
#    return False

def GDMM_Check():
    #spotdate=spotdate1.get()
    #lastdate=lastdate1.get()
    #if (spotdate=='')|(lastdate==''):
    #    msbox.showinfo(title='Notice',message='日期不能为空')
    #    return 0
    
    #try:
    #    spotdate=pd.to_datetime(spotdate,errors='raise')
    #    lastdate=pd.to_datetime(lastdate,errors='raise')
    #except Exception:
    #    msbox.showinfo(title='Notice',message='请输入正确的日期格式！如：2017-09-25')
    #    return 0
    
    
    #if check_contain_chinese(path0):
    #    msbox.showinfo(title='Notice',message='当前路径为：'+path0+'。路径名不能含中文字符，请修改路径！')
    #    return 0
    
    path1='Opt_DM\TheBegin.xlsx'
    path2='Opt_DM\TheEnd.xlsx'
    path3='Opt_DM\Collateral.xlsx'
    path4='Opt_DM\Report3.xlsx'
    path5='Opt_DM\Report_Opt_Book.xlsx'
    
    
    try:
        lastmonth=pd.read_excel(path0+path1,0,encoding='gbk',keep_default_na=False)
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path1)
        return 0
    
    try:
        thismonth=pd.read_excel(path0+path2,0,encoding='gbk',keep_default_na=False)
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path2)
        return 0
    
    try:
        collateral=pd.read_excel(path0+path3,0,encoding='gbk',keep_default_na=False)
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path3)
        return 0
    
    try:
        lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl,collateral_newtype=GDMM.Check0(lastmonth,thismonth,collateral)
    except:
        msbox.showinfo(title='Notice',message='在执行第一步检查时,发生错误')
        return 0
    
    try:
        lastmonth,thismonth,collateral=GDMM.Check1(lastmonth,thismonth,collateral)
    except:
        msbox.showinfo(title='Notice',message='在执行第二步检查时,发生错误')
        return 0
    
    try:
        Outputs=GDMM.Check2(lastmonth,thismonth,collateral)
    except:
        msbox.showinfo(title='Notice',message='在执行第三步检查时,发生错误')
        return 0
    
    
    try:
        wbw1=pd.ExcelWriter(path0+path4)
        lastmonth.to_excel(wbw1,'期初表',index=False)
        thismonth.to_excel(wbw1,'期末表',index=False)
        collateral.to_excel(wbw1,'现金流',index=False)
        Outputs.to_excel(wbw1,'检查结果',index=False)
    
        if len(lastmonth_dupl)!=0:
            lastmonth_dupl.to_excel(wbw1,'期初表重复',index=False)
        if len(thismonth_dupl)!=0:
            thismonth_dupl.to_excel(wbw1,'期末表重复',index=False)
        if len(collateral_dupl)!=0:
            collateral_dupl.to_excel(wbw1,'现金流重复',index=False)
        if len(collateral_newtype)!=0:
            collateral_newtype.to_excel(wbw1,'新现金流类型',index=False)
        
        wbw1.save()
        msbox.showinfo(title='Notice',message='完成检查！输出检查报告到'+path0+path4)
    except:
        msbox.showinfo(title='Notice',message='输出报告时发生错误')
        return 0

    
    try:
        Outputs2=Outputs[['ContractID','期初表期权标的','期初表标的类型','期末表期权标的','期末表标的类型','资金表期权标的','资金表标的类型','期初表Upfront','期末表Upfront','前端支付','前端期权费','展期期权费','到期结算','部分赎回','全部赎回','期间结算','红利支付','定结期间结算']]
        val0=lastmonth[['ContractID','合约估值(结算币种)']]
        Outputs2=pd.merge(Outputs2,val0,how='left',on='ContractID')
        Outputs2=Outputs2.rename(columns={'合约估值(结算币种)':'期初表合约估值CNY'})
        
        val1=thismonth[['ContractID','合约估值(结算币种)']]
        Outputs2=pd.merge(Outputs2,val1,how='left',on='ContractID')
        Outputs2=Outputs2.rename(columns={'合约估值(结算币种)':'期末表合约估值CNY'})
        Outputs2['期初表合约估值CNY']=pd.to_numeric(Outputs2['期初表合约估值CNY'])
        Outputs2['期末表合约估值CNY']=pd.to_numeric(Outputs2['期末表合约估值CNY'])
        
        Outputs2['实盈']=Outputs2['前端支付'].replace(np.nan,0)+Outputs2['前端期权费'].replace(np.nan,0)+Outputs['展期期权费'].replace(np.nan,0)+Outputs['到期结算'].replace(np.nan,0)+Outputs['部分赎回'].replace(np.nan,0)+Outputs['全部赎回'].replace(np.nan,0)+Outputs['期间结算'].replace(np.nan,0)+Outputs['红利支付'].replace(np.nan,0)+Outputs['定结期间结算'].replace(np.nan,0)+Outputs['期初表Upfront'].replace(np.nan,0)-Outputs['期末表Upfront'].replace(np.nan,0)
        Outputs2['公允价值变动损益']=Outputs2['期末表合约估值CNY'].replace(np.nan,0)-Outputs2['期初表合约估值CNY'].replace(np.nan,0)+Outputs2['期末表Upfront'].replace(np.nan,0)-Outputs2['期初表Upfront'].replace(np.nan,0)
        
        wbw2=pd.ExcelWriter(path0+path5)
        lastmonth.to_excel(wbw2,'期初表',index=False)
        thismonth.to_excel(wbw2,'期末表',index=False)
        collateral.to_excel(wbw2,'现金流',index=False)
        Outputs2.to_excel(wbw2,'台账',index=False)
        wbw2.save()
        msbox.showinfo(title='Notice',message='生成期权台账,输出到:'+path0+path5)
    except:
        msbox.showinfo(title='Notice',message='生成期权台账时发生错误，但不影响数据检查报表')
        return 0




class APP_class():
    def __init__(self,root):
        self.root=root
        self.setupUI1()
        
    def setupUI1(self):
        root=self.root
        tk.Label(root,text='期初日期:').grid(row=0,column=0,sticky="w")
        tk.Label(root,text='期末日期:').grid(row=0,column=2,sticky="w")
        
        spotdate1=tk.StringVar()
        Entry1=tk.Entry(root,textvariable=spotdate1,width=12)
        Entry1.grid(row=0,column=3,sticky="e")
        
        lastdate1=tk.StringVar()
        Entry2=tk.Entry(root,textvariable=lastdate1,width=12)
        Entry2.grid(row=0,column=1,sticky="e")
        
        tk.Button(root,text="检查数据",command=lambda:GDMM_Check(),width=10).grid(row=1,column=0,sticky="e")
    
    

if __name__=="__main__":
    root=tk.Tk()
    root.title('iAccountingTool.0.0.1')
    APP1=APP_class(root)
    
    root.mainloop()