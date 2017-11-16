# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 21:51:42 2017

@author: Frank

ref:https://github.com/pdfminer/pdfminer.six/blob/master/docs/programming.html
    http://blog.csdn.net/u013421629/article/details/72764737?locationNum=2&fps=1
"""

# coding=utf-8
#import sys
#reload(sys)
#sys.setdefaultencoding('utf-8')
import time
time1=time.time()
import os.path
from pdfminer.pdfparser import PDFParser#,PDFDocument
from pdfminer.pdfdocument import PDFDocument,PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfpage import PDFPage
#from pdfminer.pdfinterp import PDFTextExtractionNotAllowed

result=[]
class CPdf2TxtManager():
    def __init__(self,filepath=None,password=''):
        self.initialize()
        if filepath is not None:
            self.open(filepath,password)
        
    def initialize(self):
        self.isopened = False
        self.file = None
        self.parser = None
        self.document = None
        self.rsmgr = PDFResourceManager()
        self.laparams = LAParams()
        self.device = PDFPageAggregator(self.rsmgr, laparams=self.laparams)
        self.interpreter = PDFPageInterpreter(self.rsmgr, self.device)

    def open(self,filename,password=''):
        if not self.isopened:
            self.file = open(filename, 'rb') # 以二进制读模式打开
            self.parser = PDFParser(self.file)
            self.document = PDFDocument(self.parser, password)
            self.parser.set_document(self.document)
            self.isopened = True
        
    def close(self):
        if self.isopened:
            self.file.close()
            self.isopened = True
        
    def changePdfToText(self, filePath, password=''):
        self.open(filePath,password)
        # 检测文档是否提供txt转换，不提供就忽略
        if not self.document.is_extractable:
            raise PDFTextExtractionNotAllowed

        fileNames = os.path.splitext(filePath)
        # 循环遍历列表，每次处理一个page的内容
        with open(fileNames[0] + '.txt','wb') as f:
            for page in PDFPage.create_pages(self.document):#for page in doc.get_pages(): # doc.get_pages() 获取page列表
                self.interpreter.process_page(page)
                # 接受该页面的LTPage对象
                layout = self.device.get_result()
                for x in layout:
                    if hasattr(x, "get_text"):
                        # print x.get_text()
                        results = x.get_text()
                        #result.append(x.get_text())
                        #print(results)
                        if isinstance(results, str):
                            results += '\n'
                            results = results.encode('utf-8')
                        f.write(results)
        self.close()


    def getpagenum(self):
        pagenum=0
        for page in PDFPage.create_pages(self.document):
            pagenum += 1
        self.close()
        print(pagenum)
        return pagenum
        
inpath = os.path.join(os.getcwd(),'..\\w.pdf')

if __name__ == '__main__':
    '''''
     解析pdf 文本，保存到txt文件中
    '''

    pdf2TxtManager = CPdf2TxtManager(inpath)
    pdf2TxtManager.getpagenum()

    # print result[0]
    time2 = time.time()

    print(u'ok,解析pdf结束!')
    print(u'总共耗时：' + str(time2 - time1) + 's')