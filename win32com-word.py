# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 16:00:31 2017

@author: Frank
#ref:http://new.galalaly.me/2011/09/use-python-to-parse-microsoft-word-documents-using-pywin32-library/
    https://stackoverflow.com/questions/10366596/how-to-read-contents-of-an-table-in-ms-word-file-using-python
    https://stackoverflow.com/questions/36193159/page-number-python-docx
    
"""

# -*- coding: cp936 -*-
#导入随机数模块
import os
import glob
import random
#导入win32com模块,用来操作Word
#import win32com
from win32com.client import Dispatch, constants
WORD = 'Word.Application'

def writedocument(doc):
    #准备对文档头部进行操作
    myRange = doc.Range(0,0)#从第0行第0个字开始：
    myRange.Style.Font.Name = "宋体"#设置字体
    myRange.Style.Font.Size = "16"#设置为三号
    #========以下为文章的内容部分=======
    #文章标题（用\n来控制文字的换行操作）
    title='XXXXX会\n会议时间: '
    #会议时间
    timelist=['1月9日','1月16日','1月23日','1月30日',
     '2月6日','2月13日','2月20日','2月27日',
     '3月6日','3月13日','3月20日','3月27日',
     '4月3日','4月10日','4月17日','4月24日',
     '5月8日','5月15日','5月22日','5月29日',
     '6月5日','6月12日','6月19日','6月26日',
     '7月3日','7月10日','7月17日','7月24日'
     ]
    #会议地点
    addrlist=['\n会议地点: 地点AXXX\n主持人: 张X\n',
     '\n会议地点: 地点BXXXX主持人: 吴X\n',
     '\n会议地点: 地点CXXXX\n主持人: 王X\n',
     '\n会议地点: 地点DXXXX\n主持人: 冉X\n',
     '\n会议地点: 地点EXXXX\n主持人: 李X\n',
     ]
    #参加人员
    member='参加人员: XXX,XXX,XXX,XXX,XXX,XXX,XXX。\n会议内容：\n '
    #四段文字
    list1=['第一段(A型)\n','第一段(B型)\n','第一段(C型)\n','第一段(D型)\n']
    list2=['第二段(A型)\n','第二段(B型)\n','第二段(C型)\n','第二段(D型)\n']
    list3=['第三段(A型)\n','第三段(B型)\n','第三段(C型)\n','第三段(D型)\n']
    list4=['第四段(A型)\n','第四段(B型)\n','第四段(C型)\n','第四段(D型)\n']
    #开始循环操作，往Word里面写文字
     #先开始遍历地点（A,B,C,D,E四个地区）
    str3='dk'
    for addr in addrlist:
        #遍历28个日期
        for time in timelist:
        #随机生成四个数（范围0-3）
            aa=random.randint(0,3)
            bb=random.randint(0,3)
            cc=random.randint(0,3)
            dd=random.randint(0,3)
            #从文件开头依次插入标题、时间、地点、人物
            myRange.InsertAfter(title)
            myRange.InsertAfter(time)
            myRange.InsertAfter(addr)
            myRange.InsertAfter(str3)
            #在后面继续添加随机选取的四段文字
            myRange.InsertAfter(list1[aa])
            myRange.InsertAfter(list2[bb])
            myRange.InsertAfter(list3[cc])
            myRange.InsertAfter(list4[dd])
            
def writedocument_simple(doc):
    #准备对文档头部进行操作
    myRange = doc.Range(0,0)#从第0行第0个字开始：
    myRange.Style.Font.Name = "宋体"#设置字体
    myRange.Style.Font.Size = "16"#设置为三号
    for text in ['aa','bb','cc']:
        myRange.InsertAfter(text)
            
class Word:
    
    def __init__(self):
        self.app = Dispatch(WORD)#win32.gencache.EnsureDispatch(WORD)
        self.app.Visible = 0#0表示在后台操作。设为1则在前端能看到Word界面。
        self.app.DisplayAlerts = 0#不显示警告
        self.doc = None
        
    def open(self, filename):
        if self.doc is None:
            self.doc = self.app.Documents.Open(filename)
            #self.doc = self.app.Documents.Add()
    
    def open_multi(self):
        for file in glob.glob( os.path.join('', '*.docx') ):
            fullpath = os.path.join(os.getcwd(),file)
            self.doc = self.app.Documents.Open(fullpath)
            self.doc = self.app.Documents.Add()#?
            if not self.doc.CheckGrammar:
                print("Did not pass the grammar and spelling check")

    def save(self,filename,SaveChanges=False):
        self.doc.SaveAs(filename)
    
    def getparagraphs(self):
        doc = self.app.ActiveDocument
        for para in doc.Paragraphs:
            #print(para)
            range = para.Range
            print('%3d %s' % (range.Words.Count, range.Text[0:6]))

    @staticmethod
    def doc2pdf(doc,filename):
        doc.ExportAsFixedFormat(filename,\
                                constants.wdExportFormatPDF,\
                                Item = constants.wdExportDocumentWithMarkup,\
                                CreateBookmarks = constants.wdExportCreateHeadingBookmarks)
        
    def checkstyle(self,savefilename):
        word = Word()
        word.open_multi()
        fonts = []
        sizes = []
        styles = []
        effect = False
        # For every word in the document
        for word_t in self.doc.Words:
             if not word_t.Font.Name in fonts:
                  fonts.append(word_t.Font.Name)
             if not word_t.Font.Size in sizes:
                  sizes.append(word_t.Font.Size)
             if not word_t.Style in styles:
                  styles.append(word_t.Style)
             if word_t.Font.Bold or word_t.Font.DoubleStrikeThrough or word_t.Font.Emboss or word_t.Font.Italic or word_t.Font.Underline or word_t.Font.Engrave or word_t.Font.Shadow or word_t.Font.Shading or word_t.Font.StrikeThrough or word_t.Font.Subscript or word_t.Font.Superscript or word_t.Font.SmallCaps or word_t.Font.AllCaps:
                  effect = True
                  
    def replace(self, source, target):
        self.app.Selection.HomeKey(Unit=constants.wdLine)
        find = self.app.Selection.Find
        find.Text = "%" + source + "%"
        self.app.Selection.Find.Execute()
        self.app.Selection.TypeText(Text=target)

    @staticmethod
    def search_replace(file, find_str, replace_str):
        word = Word()
        word.open(file)
        word.app.Selection.Find.Text = find_str
        found = word.app.Selection.Find.Execute()
        doc = word.app.ActiveDocument
        if found:
            word.app.Selection.TypeText(replace_str)
            doc.Close(SaveChanges=True)
        else:
            doc.Close(SaveChanges=False)
        return found
    
    @staticmethod
    def replaceall(file, find_str, replace_str):
        word = Word()
        word.open(file)
        wdFindContinue = 1
        wdReplaceAll = 2
        word.app.Selection.Find.Execute(find_str, False, False, False, False, False, \
                                   True, wdFindContinue, False, replace_str, wdReplaceAll)
        word.close(SaveChanges=True)
        
    @staticmethod
    def createdocument(filename):
        word = Word()        
        doc = word.app.Documents.Add()
        writedocument_simple(doc)
        doc.SaveAs(filename)
        doc.Close()
        
    def printdoc(self):
        self.app.Application.PrintOut()
        
    def close(self,SaveChanges=False):
        self.app.ActiveDocument.Close(SaveChanges=SaveChanges)
        
    def exit(self):
        self.app.Quit()

def print_testdrive(customer, vehicle, appointment):
    word = Word()
    word.open(r"test.doc")
    if customer.nameprefix:
        word.replace("nameprefix",customer.nameprefix)
    else:
        word.replace("nameprefix", " ")
    if customer.firstname:
        word.replace("firstname",customer.firstname)
    else:
        word.replace("firstname", " ")
    if customer.middlename:
        word.replace("middlename",customer.middlename)
    else:
        word.replace("middlename", " ")
    if customer.lastname:
        word.replace("lastname",customer.lastname)
    else:
        word.replace("lastname", " ")
    if customer.namesuffix:
        word.replace("namesuffix",customer.namesuffix)
    else:
        word.replace("namesuffix", " ")
    word.replace("time",appointment.time)
    word.replace("date",appointment.date)
    word.replace("reservation",appointment.reservation)
    word.replace("VIN",vehicle.VIN)
    word.replace("Make",vehicle.Make)
    word.replace("Model",vehicle.Model)
    word.replace("Color",vehicle.Color)
    word.replace("OptionPackage",vehicle.OptionPackage)
    word.replace("Accessories",vehicle.Accessories)
    word.replace("ModelYear",vehicle.ModelYear)
    word.printdoc()
    word.close()

filepath = os.path.join(os.getcwd(),'test.docx')
outpath = os.path.join(os.getcwd(),'out.docx')

def testcreate():
    Word.createdocument(outpath)
    
def main():
    word = Word()
    word.open(filepath)
    word.getparagraphs()
    
if __name__ == "__main__":
    main()