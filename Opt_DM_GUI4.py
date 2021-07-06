# -*- coding: utf-8 -*-
"""
Created on Fri Feb 23 11:51:15 2018

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
import Opt_DM_GUI3
from Opt_DM_GUI2 import GenNote
from Opt_DM_GUI5 import GenBook

try:
    #path0=os.path.dirname(os.path.realpath(__file__))+'\\' 
    path0=os.getcwd()+'\\' 
except:
    #path0=os.path.dirname((os.path.realpath(__file__)).decode('gb2312'))+'\\'
    path0=(os.getcwd()).decode('gb2312')+'\\' 
    

class APP_class(Opt_DM_GUI3.APP_class):
    def __init__(self,root):
        self.root=root
        self.setupUI2()
        self.setupUI3()
    
    def setupUI2(self):
        self.setupUI1()
        tk.Button(self.root,text="生成凭证",command=GenNote,width=10).grid(row=1,column=1,sticky="e")
    
    def setupUI3(self):
        self.setupUI1()
        tk.Button(self.root,text="生成台账(按标的)",command=GenBook,width=18).grid(row=1,column=2,sticky="e")



if __name__=="__main__":
    root=tk.Tk()
    root.title('iAccountingTool.0.1.1')
    APP1=APP_class(root)
    root.mainloop()