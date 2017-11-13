# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 09:30:33 2017

@author: Frank

ref:https://github.com/mikemaccana/python-docx/blob/master/example-makedocument.py
    https://stackoverflow.com/questions/25228106/how-to-extract-text-from-an-existing-docx-file-using-python-docx
    https://stackoverflow.com/questions/22765313/when-import-docx-in-python3-3-i-have-error-importerror-no-module-named-excepti
"""

import docx

class DOC:
    
    def __init__(self,filename=None):
        self.initialize()
        if filename is not None:
            self.load(filename)
        
    def load(self,filename):
        self.doc = docx.Document(filename)
        
    def initialize(self):
        self.doc = None
    
    def gettext(self):
        fullText = []
        for i,para in enumerate(self.doc.paragraphs):
            print(i,para.text)
            fullText.append(para.text)
        return '\n'.join(fullText)

import os

curdir = os.getcwd()
filename = 'test.docx'#'blank.docx'#
filepath_in = os.path.join(curdir,filename)
filepath_out = os.path.join(curdir,'out.docx')

def main():
    doc = DOC(filepath_in)
    doc.gettext()
    
if __name__ == "__main__":
    main()