# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 14:38:15 2017

@author: Frank

explanation:OpenXML 
    1. Big Frame
    <w:document>
        <w:body>
            <w:p>#paragraph</w:p>
            <w:sectPr>#</w:sectPr>
        </w:body>
    </w:document>
    2. Small Frame
    - Paragraph
    <w:p w:rsidR="00450526" w:rsidRDefault="00450526" w:rsidP="00450526">
        <w:pPr>
            <w:pStyle w:val="Heading1"/>
        </w:pPr>
        <w:r>#content
            <w:t>TEXT</w:t>
        </w:r>
    </w:p>
    
    - Text Part
    <w:r>#content
        <w:rPr>
            <w:b/>
        </w:rPr>
        <w:t>TEXT</w:t>
    </w:r>
    
    - Setting
    
"""

from zipfile import ZipFile
from lxml import etree

def opendocx(docx_filename):
    with open(docx_filename,'r+b') as f:
        zip = ZipFile(f)
        xml_content = zip.read('word/document.xml')
        return etree.fromstring(xml_content)
    return None
  
class OpenXML:
    
    def __init__(self,zbdoc,tbdoc=None,tbdir=None,tbsqlite=None):
        self.recruit = zbdoc#zb stands for 招标
        self.output = tbdoc#tb stands for 投标
        self.system = []#main frame for output document
        self.db_img = tbdir
        self.db_sqlite = tbsqlite
        self.setkeyword = None
        pass
    
    def open(self,filename):
        pass
    
    def save(self,filename):
        pass
    
    def copypaste(self):
        pass
    
    def setkeyword(self):
        pass
    
    def findpage(self):
        pass
    
    def findcoverpage(self):
        #keyword:投标文件格式，
        #        项目名称，
        #        投标文件，
        #        投标人，
        #        盖章单位，
        #        __年__月__日
        pass
    
    def findindexpage(self):
        #keyword:投标函，
        #        承诺书
        #        法定代表人（法人）身份证明
        #        授权委托书
        #        投标保证金
        #        已标价工程量清单
        #        一，二，三，四，五，六，七，八，九，十
        pass
