# -*- coding: utf-8 -*-
"""
Created on Sat Nov 11 10:11:47 2017

@author: Frank

ref:https://automatetheboringstuff.com/chapter13/
    https://github.com/mstamy2/PyPDF2
"""

import sys
# sys.setdefaultencoding() does not exist, here!
#reload(sys) # Reload does the trick!
#sys.setdefaultencoding('UTF8')
import PyPDF2
from PyPDF2.pdf import ContentStream
from PyPDF2.utils import b_,u_
from PyPDF2.generic import TextStringObject
import codecs

def byte2text(byte):
    return repr(str(byte))

def text2byte(text):
    return bytes(text, 'utf-8')
    
class PDF:
    
    def __init__(self,filename=None):
        self.initialize()
        if filename is not None:
            self.load(filename)
 
    def initialize(self):
        self.pdfreader = None
        self.pdfwriter = PyPDF2.PdfFileWriter()
        self.pageNum = 0
        self.filehandle = None
        
    def load(self,filename):
        self.filehandle = open(filename,'r+b')
        self.pdfreader = PyPDF2.PdfFileReader(self.filehandle)
        self.pageNum = self.pdfreader.numPages
        
    def close(self):
        self.filehandle.close()
        
    def save(self,filename):
        with open(filename, 'wb') as f:
            self.pdfwriter.write(f)
    
    def getpage(self,pagenum):
        return self.pdfreader.getPage(pagenum)
    
    def addpage(self,page):
        self.pdfwriter.addPage(page)
        
    def extracttext(self,pagenum):
        page = self.pdfreader.getPage(pagenum)
        content = page.getContents()
        if not isinstance(content, ContentStream):
            content = ContentStream(content, page.pdf)
        text = u_("")
        for operands, operator in content.operations:
            #print(operator)
            if operator == b_("Tj"):
                _text = operands[0]
                #print(byte2text(_text))
                print(_text)
                #if isinstance(_text, TextStringObject):
                #    text += _text
                    #print(_text.decode('utf-8'))
        
    def decrypt(self,pw):
        self.pdfReader.decrypt(pw)
 
    def encrypt(self,pw):
        self.pdfwriter.encrypt(pw)
        
    def isEncrypted(self):
        return self.pdfreader.isEncrypted
    
    def getpagenum(self):
        return self.pdfreader.numPages

if __name__ == "__main__":
    pdf = PDF('w.pdf')
    pdf.extracttext(3)
    pdf.close()
    
    
