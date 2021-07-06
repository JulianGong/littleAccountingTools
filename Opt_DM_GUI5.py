# -*- coding: utf-8 -*-
"""
Created on Wed Mar 27 14:29:46 2019

@author: 028375
"""

from __future__ import unicode_literals, division

try:
    import tkinter as tk
    import tkinter.messagebox as msbox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as msbox
    
import os, copy
import pandas as pd
import numpy as np
import Opt_DM_GUI3

try:
    #path0=os.path.dirname(os.path.realpath(__file__))+'\\' 
    path0=os.getcwd()+'\\' 
except:
    #path0=os.path.dirname((os.path.realpath(__file__)).decode('gb2312'))+'\\'
    path0=(os.getcwd()).decode('gb2312')+'\\' 
    
path5='Opt_DM\Report_Opt_Book.xlsx'
path6='Opt_DM\TheAssess.xlsx'
path7='Opt_DM\Report_Opt_Book2.xlsx'


def GenBook():
    
    try:
        Outputs2=pd.read_excel(path0+path5,3,encoding='gbk',keep_default_na=False)
        
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path5)
        return 0
    
    try:
        TheAssess=pd.read_excel(path0+path6,0,encoding='gbk',keep_default_na=False)
        
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path6)
        return 0
    
    try:
        Outputs2['期权标的']=Outputs2['期初表期权标的'].copy()
        Outputs2.loc[Outputs2['期权标的']=='',['期权标的']]=Outputs2[Outputs2['期权标的']=='']['期末表期权标的']
        Outputs2.loc[Outputs2['期权标的']=='',['期权标的']]=Outputs2[Outputs2['期权标的']=='']['资金表期权标的']
        Outputs2['期权标的']=Outputs2['期权标的'].str.rstrip(';,,,；,.')
        
        Outputs2['标的类型']=Outputs2['期初表标的类型'].copy()
        Outputs2.loc[Outputs2['标的类型']=='',['标的类型']]=Outputs2[Outputs2['标的类型']=='']['期末表标的类型']
        Outputs2.loc[Outputs2['标的类型']=='',['标的类型']]=Outputs2[Outputs2['标的类型']=='']['资金表标的类型']
        
            
        Outputs3=Outputs2.groupby(['期权标的','标的类型'])['实盈','公允价值变动损益'].sum().reset_index()
        
    except BaseException:
        msbox.showinfo(title='Notice',message='处理期权端数据时发生错误！')
        return 0
    
    
    try:
        TheAssess=TheAssess[(TheAssess['部门2']=='股权衍生品业务线') & (TheAssess['业务类型']=='场外期权')]
        TheAssess['投资收益']=TheAssess['卖出利润']-TheAssess['交易费用']
        TheAssess['分红']=TheAssess['本日计提利息']
        TheAssess['浮动盈亏2']=pd.to_numeric(TheAssess['计提公允价值损益'])
        TheAssess2=TheAssess[['代码','名称', '市场','金融分类', '投资收益','分红','浮动盈亏2']].reset_index(drop=True)
        TheAssess2.loc[TheAssess2['市场']=='X_SZSC',['市场']]='X_SHSC' #部分港股通市场串号
        
        TheAssess2=TheAssess2.groupby(['代码','名称','市场','金融分类'])['投资收益','分红','浮动盈亏2'].sum().reset_index()
        TheAssess2['修正后的代码']=TheAssess2['代码'].copy()
        
        
        for i in list(range(len(TheAssess2))):
            try:
                if TheAssess2['市场'][i]=='XSHE':
                    TheAssess2['修正后的代码'][i]='{0:0>6}'.format(TheAssess2['代码'][i])+'.SZ'
                elif TheAssess2['市场'][i]=='XSHG':
                    TheAssess2['修正后的代码'][i]='{0:0>6}'.format(TheAssess2['代码'][i])+'.SH'
                elif (TheAssess2['市场'][i]=='X_SZSC')|(TheAssess2['市场'][i]=='X_SHSC'):
                    TheAssess2['修正后的代码'][i]='{0:0>4}'.format(TheAssess2['代码'][i])+'.HK'
                elif TheAssess2['市场'][i]=='X_CNFFEX':
                    if TheAssess2['名称'][i] in (['螺纹钢','热轧卷板','锌','铝','黄金','白银','镍','锡','阴极铜','纸浆','石油沥青','天然橡胶']):
                        TheAssess2['修正后的代码'][i]=(TheAssess2['代码'][i]).upper()+'.SHF'
                    elif TheAssess2['名称'][i] in (['原油']):
                        TheAssess2['修正后的代码'][i]=(TheAssess2['代码'][i]).upper()+'.INE'
                    elif TheAssess2['名称'][i] in (['聚氯乙烯','聚丙烯','大豆原油','铁矿石','冶金焦炭','豆粕','棕榈油','玉米','线型低密度聚乙烯','冶金焦炭']):
                        TheAssess2['修正后的代码'][i]=(TheAssess2['代码'][i]).upper()+'.DCE'
                    elif TheAssess2['名称'][i] in (['动力煤','鲜苹果','一号棉花','精对苯二甲酸(PTA)','甲醇','白砂糖']):
                        TheAssess2['修正后的代码'][i]=(TheAssess2['代码'][i]).upper()+'.CZC'   
                        
                        
            except: 0
            
    except BaseException:
        msbox.showinfo(title='Notice',message='处理现货端数据时发生错误！')
        return 0
        
    
    try:
        TheAssess3=TheAssess2.rename(columns={'修正后的代码':'期权标的'})
        
        Outputs3=Outputs3[Outputs3['期权标的']!='']
        TheAssess3=TheAssess3[TheAssess3['期权标的']!='']
        
        Outputs5=pd.merge(Outputs3,TheAssess3,how='outer',on='期权标的')
        
        wbw3=pd.ExcelWriter(path0+path7)
        Outputs2.to_excel(wbw3,'期权台账',index=False)
        TheAssess.to_excel(wbw3,'考核报表',index=False)
        Outputs3.to_excel(wbw3,'期权台账_按标的汇总',index=False)
        TheAssess2.to_excel(wbw3,'考核报表_按标的汇总',index=False)
        Outputs5.to_excel(wbw3,'对冲结果',index=False)
        
        wbw3.save()
        msbox.showinfo(title='Notice',message='生成期权台账(按标的),输出到:'+path0+path7)
    except:
        msbox.showinfo(title='Notice',message='生成期权台账(按标的)发生错误！')
        return 0




class APP_class(Opt_DM_GUI3.APP_class):
    def __init__(self,root):
        self.root=root
        self.setupUI2()
    
    def setupUI2(self):
        self.setupUI1()
        tk.Button(self.root,text="生成台账(按标的)",command=GenBook,width=18).grid(row=1,column=1,sticky="e")
        



if __name__=="__main__":
    root=tk.Tk()
    root.title('iAccountingTool.0.1.1')
    APP1=APP_class(root)
    root.mainloop()
    