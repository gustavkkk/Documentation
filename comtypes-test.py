# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 21:23:50 2017

@author: Frank
"""

#import sys
import os
import comtypes.client

wdFormatPDF = 17
WORD = 'Word.Application'

class Word:
    
    def __init__(self,filename=None):
        self.initialize()
        if filename is not None:
            self.load(filename)
    
    def initialize(self):
        self.word = comtypes.client.CreateObject(WORD)
        self.word.Visible = False
        self.doc = None
        
    def load(self,filename):
        self.doc = self.word.Documents.Open(filename)
    
    def savePDF(self,filename):
        self.doc.SaveAs(filename, FileFormat=wdFormatPDF)
        
    @staticmethod
    def doc2pdf(fn_in,fn_out):
        word = Word(fn_in)
        word.savePDF(fn_out)
        word.close()
        
    def close(self):
        self.doc.Close()
        self.word.Quit()

inpath = os.path.join(os.getcwd(),'test.docx')
outpath = os.path.join(os.getcwd(),'test.pdf')

def main():
    Word.doc2pdf(inpath,outpath)
    
if __name__ == "__main__":
    main()
    