# -*- coding: utf-8 -*-
"""
Created on Thu Nov 09 14:37:17 2017

@author: 028375
"""
from __future__ import unicode_literals, division

try:
    import tkinter as tk
    import tkinter.messagebox as msbox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as msbox
    
import Opt_DM_GUI2 as OptGUI, ELN_GUI2 as ELNGUI
import Opt_DM_GUI4 as OptGUI_new
from win32com.client import Dispatch
import os
try:
    #path0=os.path.dirname(os.path.realpath(__file__))+'\\' 
    path0=os.getcwd()+'\\' 
except:
    #path0=os.path.dirname((os.path.realpath(__file__)).decode('gb2312'))+'\\'
    path0=(os.getcwd()).decode('gb2312')+'\\' 

def Gent3():
    try:
        xlApp = Dispatch('Excel.Application')
        path1=path0+'iAT(ExcelVBA) 0.3.4.xlsm'
        xlApp.Workbooks.Open(path1)
        xlApp.visible=True
    except:
        msbox.showinfo(title='Notice',message='找不到'+path1+'!')


class APP_class():
    def __init__(self,root):
        self.root=root
        self.setupUI1()
    
    def setupUI1(self):
        root=self.root
        menubar1=tk.Menu(root)
        menubar1.add_command(label='股衍境内期权',command=self.Gent1)
        menubar1.add_command(label='股衍跨境期权',command=Gent3)
        menubar1.add_command(label='股衍收益凭证',command=self.Gent2)
        menubar1.add_command(label='股衍境内期权(新)',command=self.Gent4)
        root.config(menu=menubar1)
        
    def Gent1(self):
        root1=tk.Toplevel(self.root)
        root1.grab_set()
        root1.focus_set()
        root1.wm_attributes('-topmost',1)
        
        GUI1=OptGUI.APP_class(root1)
        GUI1.setupUI2()
        
        root1.title('股衍境内期权')
        root1.geometry('300x50+400+300')
    
    def Gent2(self):
        root2=tk.Toplevel(self.root)
        root2.grab_set()
        root2.focus_set()
        root2.wm_attributes('-topmost',1)
        
        GUI2=ELNGUI.APP_class(root2)
        GUI2.setupUI2()
        
        root2.title('股衍收益凭证')
        root2.geometry('300x50+400+300')
        
    def Gent4(self):
        root4=tk.Toplevel(self.root)
        root4.grab_set()
        root4.focus_set()
        root4.wm_attributes('-topmost',1)
        
        GUI4=OptGUI_new.APP_class(root4)
        GUI4.setupUI2()
        
        root4.title('股衍境内期权(新)')
        root4.geometry('350x50+400+300')  
        

if __name__=="__main__":
    root=tk.Tk()
    root.geometry('300x80+500+200')
    
    root.title('iAccountingTool 0.2.4')
    APP1=APP_class(root)
    root.mainloop()