# -*- coding: utf-8 -*-
"""
Created on Tue Nov 7 15:14:59 2017

@author: 028375
"""

from __future__ import unicode_literals, division

try:
    import tkinter as tk
    import tkinter.messagebox as msbox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as msbox
    
import Opt_DM_Script1 as GDMM
import os
import pandas as pd


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

def GDMM_Check(spotdate1,lastdate1):
    spotdate=spotdate1.get()
    lastdate=lastdate1.get()
    if (spotdate=='')|(lastdate==''):
        msbox.showinfo(title='Notice',message='日期不能为空')
        return 0
    
    try:
        spotdate=pd.to_datetime(spotdate,errors='raise')
        lastdate=pd.to_datetime(lastdate,errors='raise')
    except Exception:
        msbox.showinfo(title='Notice',message='请输入正确的日期格式！如：2017-09-25')
        return 0
    
    
    #if check_contain_chinese(path0):
    #    msbox.showinfo(title='Notice',message='当前路径为：'+path0+'。路径名不能含中文字符，请修改路径！')
    #    return 0
    
    path1='Opt_DM\TheBegin.xlsx'
    path2='Opt_DM\TheEnd.xlsx'
    path3='Opt_DM\Collateral.xlsx'
    path4='Opt_DM\FX.xlsx'
    path5='Opt_DM\Report.xlsx'
    
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
        FXRate=pd.read_excel(path0+path4,0,encoding='gbk',keep_default_na=False)
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path4)
        return 0
    
    try:
        lastmonth,thismonth,collateral,lastmonth_dupl,thismonth_dupl,collateral_dupl=GDMM.Check0(lastmonth,thismonth,collateral)
    except:
        msbox.showinfo(title='Notice',message='在执行第一步检查时,发生错误')
        return 0
    
    try:
        lastmonth,thismonth,collateral=GDMM.Check1(lastmonth,thismonth,collateral,FXRate)
    except:
        msbox.showinfo(title='Notice',message='在执行第二步检查时,发生错误')
        return 0
    
    try:
        Outputs=GDMM.Check2(lastmonth,thismonth,collateral,lastdate,spotdate)
    except:
        msbox.showinfo(title='Notice',message='在执行第三步检查时,发生错误')
        return 0
    
    
    try:
        wbw1=pd.ExcelWriter(path0+path5)
        lastmonth.to_excel(wbw1,'期初表',index=False)
        thismonth.to_excel(wbw1,'期末表',index=False)
        collateral.to_excel(wbw1,'现金流',index=False)
        FXRate.to_excel(wbw1,'汇率',index=False)
        Outputs.to_excel(wbw1,'检查结果',index=False)
    
        if len(lastmonth_dupl)!=0:
            lastmonth_dupl.to_excel(wbw1,'期初表重复',index=False)
        if len(thismonth_dupl)!=0:
            thismonth_dupl.to_excel(wbw1,'期末表重复',index=False)
        if len(collateral_dupl)!=0:
            collateral_dupl.to_excel(wbw1,'现金流重复',index=False)
        wbw1.save()
        msbox.showinfo(title='Notice',message='完成检查！输出检查报告到'+path0+path5)
    except:
        msbox.showinfo(title='Notice',message='输出报告时发生错误')
        return 0



class APP_class():
    def __init__(self,root):
        self.root=root
        self.setupUI1()
        
    def setupUI1(self):
        root=self.root
        tk.Label(root,text='期初日期:').grid(row=0,sticky="w")
        tk.Label(root,text='期末日期:').grid(row=0,column=2,sticky="w")
        
        spotdate1=tk.StringVar()
        Entry1=tk.Entry(root,textvariable=spotdate1,width=12)
        Entry1.grid(row=0,column=3,sticky="e")
             
        lastdate1=tk.StringVar()
        Entry2=tk.Entry(root,textvariable=lastdate1,width=12)
        Entry2.grid(row=0,column=1,sticky="e")
        
        tk.Button(root,text="检查数据",command=lambda:GDMM_Check(spotdate1,lastdate1),width=10).grid(row=1,column=0,sticky="e")
    
    

if __name__=="__main__":
    root=tk.Tk()
    root.title('iAccountingTool.0.0.1')
    APP1=APP_class(root)
    
    root.mainloop()
