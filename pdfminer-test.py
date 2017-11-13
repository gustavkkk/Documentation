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
    def __init__(self):
        '''''
        Constructor
        '''

    def changePdfToText(self, filePath, password=''):
        file = open(filePath, 'rb') # 以二进制读模式打开
        #用文件对象来创建一个pdf文档分析器
        praser = PDFParser(file)
        # 创建一个PDF文档
        doc = PDFDocument(praser, password)
        # 连接分析器 与文档对象
        praser.set_document(doc)
        #doc.set_parser(praser)

        # 提供初始化密码
        # 如果没有密码 就创建一个空的字符串
        #doc.initialize()

        # 检测文档是否提供txt转换，不提供就忽略
        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed

        # 创建PDf 资源管理器 来管理共享资源
        rsrcmgr = PDFResourceManager()
        # 创建一个PDF设备对象
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # 创建一个PDF解释器对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        pdfStr = ''
        # 循环遍历列表，每次处理一个page的内容
        for page in PDFPage.create_pages(doc):#for page in doc.get_pages(): # doc.get_pages() 获取page列表
            interpreter.process_page(page)
            # 接受该页面的LTPage对象
            layout = device.get_result()
            for x in layout:
                if hasattr(x, "get_text"):
                    # print x.get_text()
                    result.append(x.get_text())
                    fileNames = os.path.splitext(filePath)
                    with open(fileNames[0] + '.txt','wb') as f:
                        results = x.get_text()
                        print(results)
                        #f.write(results + '\n')


inpath = os.path.join(os.getcwd(),'w.pdf')

if __name__ == '__main__':
    '''''
     解析pdf 文本，保存到txt文件中
    '''

    pdf2TxtManager = CPdf2TxtManager()
    pdf2TxtManager.changePdfToText(inpath)

    # print result[0]
    time2 = time.time()

    print(u'ok,解析pdf结束!')
    print(u'总共耗时：' + str(time2 - time1) + 's')