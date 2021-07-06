# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 16:09:22 2017

@author: 028375
"""

from __future__ import unicode_literals, division
try:
    import tkinter as tk
    import tkinter.messagebox as msbox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as msbox

import os
import pandas as pd
import ELN_GUI1
import ELN_Script2 as ELN2

try:
    #path0=os.path.dirname(os.path.realpath(__file__))+'\\' 
    path0=os.getcwd()+'\\' 
except:
    #path0=os.path.dirname((os.path.realpath(__file__)).decode('gb2312'))+'\\'
    path0=(os.getcwd()).decode('gb2312')+'\\' 
    

def GenNote():
    
    path1='ELN\Collateral2.xlsx'
    path2='ELN\Report2.xlsx'   
    
    
    try:
        collateral=pd.read_excel(path0+path1,0,encoding='gbk',keep_default_na=False)
    except BaseException:
        msbox.showinfo(title='Notice',message='找不到文件'+path0+path1)
        return 0
    
    try:
        collateral, collateral_drop=ELN2.check1(collateral)
    except:
        msbox.showinfo(title='Notice',message='在执行数据转换时,发生错误')
        return 0
    
    try:
        Notes=ELN2.genNote(collateral)
    except:
        msbox.showinfo(title='Notice',message='在生成凭证时,发生错误')
        return 0
    
    try:
        wbw1=pd.ExcelWriter(path0+path2)
        collateral.to_excel(wbw1,'现金流',index=False)
        Notes.to_excel(wbw1,'凭证',index=False)
        if len(collateral_drop)!=0:
            collateral_drop.to_excel(wbw1,'现金流错误',index=False)
        wbw1.save()
        msbox.showinfo(title='Notice',message='已生成凭证！输出凭证到 '+path0+path2)
    except:
        msbox.showinfo(title='Notice',message='输出凭证时发生错误')
        return 0


class APP_class(ELN_GUI1.APP_class):
    def __init__(self,root):
        self.root=root
        self.setupUI2()
    
    def setupUI2(self):
        self.setupUI1()
        tk.Button(self.root,text="生成凭证",command=GenNote,width=10).grid(row=1,column=1,sticky="e")
        



if __name__=="__main__":
    root=tk.Tk()
    root.title('iAccountingTool.0.1.1')
    APP1=APP_class(root)
    root.mainloop()
