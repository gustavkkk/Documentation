# -*- coding: utf-8 -*-
"""
Created on Sat Nov 11 10:14:32 2017

@author: Frank

ref:http://virantha.com/2013/08/16/reading-and-writing-microsoft-word-docx-files-with-python/
    https://stackoverflow.com/questions/19483775/python-zipfile-extract-doesnt-extract-all-files
    https://stackoverflow.com/questions/36193159/page-number-python-docx
    http://effbot.org/zone/element.htm
"""

from zipfile import ZipFile
#from gzip import GzipFile
from lxml import etree
import jieba
from config import dic,keywords

for word in dic:
    jieba.add_word(word)
             
def check_element_is(element, type_char):
     word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
     return element.tag == '{%s}%s' % (word_schema,type_char)

def get_text(element):
    string=''
    for node in element.iter(tag=etree.Element):
        if check_element_is(node, 't'):
            string += node.text
    return ''.join(string.split())
            
def unzip(zipfilename, dest_dir):
    with ZipFile(zipfilename) as zf:
        zf.extractall(dest_dir)
       
import tempfile,shutil

word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
para_tag = '{%s}%s' % (word_schema,'p')
table_tag = '{%s}%s' % (word_schema,'tbl')
text_tag = '{%s}%s' % (word_schema,'t')

class OpenXML:
    
    def __init__(self):
        self.file, self.zipfile, self.xml_content, self.xml_tree = None,None,None,None

    def close(self):
        self.file.close()

    def fill_in(self):
        self.fill_in_paragraph()
        self.fill_in_table()
        
    def fill_in_paragraph(self):
        for child in self.body:
            if check_element_is(child,'p'):
                text = get_text(child)
                if text == "":
                    continue
                print(text)
                
    def fill_in_table(self):
        pass
    
    def find_cover_page_by_index(self):
        # find the position of coverpage
        # in’投标文件格式‘
        # out:‘第九章’
        index = -1
        for child in self.body:
            # only paragraph
            if check_element_is(child,'p'):
                text = get_text(child)
                if text == "":
                    continue
                seglist = jieba.cut(text, cut_all=False)
                seglist = list(seglist)
                # only with keyword "投标文件格式"
                if keywords[1] in seglist:
                    index = seglist.index(keywords[1])
                    key = seglist[index-1]
                    position = seglist[index+2]
                    break
        if index == -1:
            return -1
        # toss
        avg_paras_per_page = 20
        estimated_position = int(position) * avg_paras_per_page
        
        # find paragraph index of coverpage
        # in: ‘第九章’，‘投标文件格式’
        # out: 2320(self.body[2320] is what)
        for i,child in enumerate(self.body):
             if i < estimated_position or not check_element_is(child,'p'):
                 continue
             text = get_text(child)
             if text == "":
                 continue
             seglist = jieba.cut(text, cut_all=False)
             seglist = list(seglist)
             if key in seglist and keywords[1] in seglist:
                 if abs(seglist.index(key) - seglist.index(keywords[1])) < 2:
                     #print(i,seglist)
                     return i
             
        return -1
                    
                    
    def load(self,docx_filename):
        self.file = open(docx_filename,'r+b')
        self.zipfile = ZipFile(self.file)
        self.xml_content = self.zipfile.read('word/document.xml')
        self.xml_tree = etree.fromstring(self.xml_content)
        self.reset()
            
    @property
    def paragraphs(self):
        return self.paras
    
    @property
    def tables(self):
        return self.tbls

    @staticmethod    
    def printag(root):
        for child in root:
            print(child.tag)
            for grandchild in child:
                print(grandchild.tag)

    @staticmethod
    def printxml(filename,xmlstring):
        with open(filename,'w+b') as xml:
            xml.write(xmlstring)
            
    def process(self):
        if self.xml_tree is None:
           return
        
        #self.remove_non_format_pages()
        #self.fill_in()
        
                
    def remove_non_format_pages(self):
        #
        coverpage_index = self.find_cover_page_by_index()
        for i,child in enumerate(self.body):
            if i > coverpage_index:
                break
            parent = child.getparent()
            #print(get_text(para))
            parent.remove(child)
            child = None
        #
        self.reset()

    def reset(self):
        self.root = self.xml_tree
        self.body = self.root[0]
        self.paras = []
        self.tbls = []
        self.xml_content = etree.tostring(self.xml_tree,pretty_print=True)
        #self.paras = self.body.iter('{%s}%s' % (word_schema,'p'))
        #self.tables = self.body.iter('{%s}%s' % (word_schema,'tbl'))
        for child in self.body:
            if check_element_is(child,'p'):
                self.paras.append(child)
            elif check_element_is(child,'tbl'):   
                self.tbls.append(child)
        
    def save(self,output_filename):
        #in:xml_tree
        #out:.docx
        tmp_dir = tempfile.mkdtemp()
        self.zipfile.extractall(tmp_dir)
        with open(os.path.join(tmp_dir,'word/document.xml'), 'w+b') as f:
            xmlstr = etree.tostring(self.xml_tree, pretty_print=True)
            f.write(xmlstr)
        # Get a list of all the files in the original docx zipfile
        filenames = self.zipfile.namelist()
        # Now, create the new zip file and add all the filex into the archive
        zip_copy_filename = output_filename
        with ZipFile(zip_copy_filename, "w") as docx_:
            for filename in filenames:
                docx_.write(os.path.join(tmp_dir,filename), filename)    
        # Clean up the temp dir
        shutil.rmtree(tmp_dir)
        #
        self.close()
        
import os 

curdir = os.getcwd()
filename = 'simohua.docx'#'blank.docx'#
filepath_in = os.path.join(curdir,filename)
filepath_out = os.path.join(curdir,'out.docx')

def main():
    oxml = OpenXML();
    oxml.load(filepath_in)
    OpenXML.printxml(,oxml.xml_content)
    #oxml.process()
    #oxml.save(filepath_out)
        
if __name__ == "__main__":
    main()

