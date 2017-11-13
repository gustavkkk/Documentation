# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 16:07:14 2017

@author: Frank
ref:http://pbpython.com/advanced-excel-workbooks.html
    http://pythonexcels.com/python-excel-mini-cookbook/
"""

#import win32com.client
from win32com.client import Dispatch
import win32
import time
import glob
import os
EXCEL="Excel.Application"

class Excel:
    
    def __init__(self):
        self.app = Dispatch(EXCEL)#win32.gencache.EnsureDispatch(EXCEL)
        self.app.Visible = False # otherwise excel is hidden
        self.workbooks = None
        self.worksheets = None
    # Handling APP
    def quit(self):
        self.app.Quit()
        
    def close(self):
        self.app.ActiveDocument.Close(SaveChanges=False)
        
    # Handling Workbook
        
    def createWB(self):
        self.workbooks = self.app.Workbooks.Add()
        
    def openWB_one(self,filename):
        self.workbooks = self.app.Workbooks.Open(filename)
        print("count of sheets:", self.workbooks.Sheets.Count)
        for sh in self.workbooks.Sheets:
            print(sh.Name)

    def openWB_all(self):
        for file in glob.glob( os.path.join('', '*.xlsx')):
            fullpath = os.path.join(os.getcwd(),file)
            self.workbooks = self.app.Workbooks.Open(fullpath)
            self.workbooks = self.app.Workbooks.Add()

    def saveWB(self,filename):
        self.workbooks.SaveAs(filename)
        
    def closeWB(self):
        self.workbooks.Close()
        
    # Handling Worksheet
    def addWS(self,sheetname):
        self.worksheets = self.workbooks.Worksheets.Add()
        self.worksheets.Name = sheetname
        
    def selectWB(self,sheetname):
        self.worksheets = self.workbooks.Worksheets(sheetname)

    # Select and Handling Selection
    def selectRegion(self,which='Range',where="B11:K11"):
        if which=='Range':#Range
            self.worksheets.Range(where).Select()
        elif which=='Column':#Cols
            self.worksheets.Columns("B:P").Select()
        elif which=='Row':#Row
            self.worksheets.Rows("1:11").Select()
    
    def chgRegionValue(self,what=12):
        self.app.Selection.Value = what
        
    # Handling Property and Value
    ## Alignment
    '''
    def setVerticalAlignment(self,which,what=win32.constants.xlRight):
        self.worksheets.Range(which).VerticalAlignment = what
        
    def setHorizontalAlignment(self,which,what=win32.constants.xlCenter):
        self.worksheets.Range(which).HorizontalAlignment  = what    
    '''        
    # Setting AutoFit
    def setRowAutoFit(self):
        self.worksheets.Rows.AutoFit()
        
    def setColAutoFit(self):
        self.worksheets.Columns.AutoFit()

    def setAutoFill(self,seed="A1:A2",field="A1:A10"):
        self.worksheets.Range(seed).AutoFill(self.worksheets.Range(field),win32.constants.xlFillDefault)
    
    def setNumberFormat(self,which="B1:B5",what="$###,##0.00"):
        self.worksheets.Range(which).NumberFormat = what
        
    # Setting Height and Width
    def setRowHeight(self,which,what,entire=False):
        if entire:
            self.worksheets.Rows(which).RowHeight  = what
        else:
            self.worksheets.Range(which).RowHeight = what
        
    def setColumnWidth(self,which,what,entire=False):
        if entire:
            self.worksheets.Columns(which).ColumnWidth  = what
        else:
            self.worksheets.Range(which).ColumnWidth  = what

    # Setting Font
    def setFont(self,which,fontname="Arial",fontsize=12):
        self.worksheets.Range(which).Font.Name = fontname
        self.worksheets.Range(which).Font.Size = fontsize
        
    # Setting Value
    def setValue_Range(self,which="A1:J10",what=21):
        self.worksheets.Range(which).Value = what
        
    def setValue_Cell(self,y=1,x=1,value=10001):
        self.worksheets.Cells(y,x).Value = value
        
    def setValue_Offset(self,y=1,x=1,offsety=2,offsetx=2,value=30001):
        self.worksheets.Cells(y,x).Offset(offsety,offsetx).Value = value
    
    # Copy and Paste
    def copyWS2WS(self,which="Sheet1",what="A1:J10"):
        ws = self.worksheets.Worksheets(which)
        ws.Range(what).Formula = "=row()*column()"
        self.workbooks.Worksheets.FillAcrossSheets(self.workbooks.Worksheets(which).Range(what))

    def setColor(self,y=1,x=1,colrindex=2):
        colrindex %= 20
        colrindex += 1
        self.worksheets.Cells(y,x).Interior.ColorIndex = colrindex
    # in Test        
    def GetValuesByCells(self):
        s = self.app.ActiveWorkbook.Sheets(1)
        startTime = time.time()
        vals = [s.Cells(r,1).Value for r in range(1,2001)]
        print(time.time() - startTime)
        return vals
    
    def GetValuesByRange(self):
        startTime = time.time()
        s = self.app.ActiveWorkbook.Sheets(1)
        vals = [v[0] for v in s.Range('A1:A2000').Value]
        print(time.time() - startTime)
        return vals


filepath_in = os.path.join(os.getcwd(),'test.xlsx')
filepath_out = os.path.join(os.getcwd(),'out.xlsx')

def TestRead():
    excel = Excel()
    excel.openWB_one(filepath_in)
    excel.closeWB()
    excel.quit()

def TestSetValue():
    excel = Excel()
    excel.createWB()
    excel.addWS("Good")
    excel.setValue_Range()
    excel.saveWB(filepath_out)
    excel.closeWB()
    excel.quit()

def TestSelect():
    excel = Excel()
    excel.createWB()
    excel.addWS("Good")
    excel.selectRegion()
    excel.chgRegionValue(22)
    excel.saveWB(filepath_out)
    excel.closeWB()
    excel.quit()   
#import xlrd
import os

if __name__ == "__main__":
    #print(getcurrentdir())
    TestSelect()
    
    