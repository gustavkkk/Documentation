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

def get_word_xml(docx_filename):
   with open(docx_filename,'r+b') as f:
      zip = ZipFile(f)
      xml_content = zip.read('word/document.xml')
   return xml_content

from lxml import etree

def get_xml_tree(xml_string):
   return etree.fromstring(xml_string)

def printxmltree(xmltree):
    print(etree.tostring(xmltree, pretty_print=True))

def _itertext(my_etree):
     """Iterator to go through xml tree's text nodes"""
     for node in my_etree.iter(tag=etree.Element):
         if _check_element_is(node, 't'):
             yield (node, node.text)

def _check_element_is(element, type_char):
     word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
     return element.tag == '{%s}%s' % (word_schema,type_char)

def _join_tags(my_etree):
    chars = []
    openbrac = False
    inside_openbrac_node = False

    for node,text in _itertext(my_etree):
        # Scan through every node with text
        for i,c in enumerate(text):
            # Go through each node's text character by character
            print(c)
            if c == '[':
                openbrac = True # Within a tag
                inside_openbrac_node = True # Tag was opened in this node
                openbrac_node = node # Save ptr to open bracket containing node
                chars = []
                print("openingbracket")
            elif c== ']':
                print("closingbracket")
                assert openbrac
                if inside_openbrac_node:
                    # Open and close inside same node, no need to do anything
                    pass
                else:
                    # Open bracket in earlier node, now it's closed
                    # So append all the chars we've encountered since the openbrac_node '['
                    # to the openbrac_node
                    chars.append(']')
                    openbrac_node.text += ''.join(chars)
                    # Also, don't forget to remove the characters seen so far from current node
                    node.text = text[i+1:]
                openbrac = False
                inside_openbrac_node = False
            else:
                # Normal text character
                if openbrac and inside_openbrac_node:
                    # No need to copy text
                    pass
                elif openbrac and not inside_openbrac_node:
                    chars.append(c)
                else:
                    # outside of a open/close
                    pass
        if openbrac and not inside_openbrac_node:
            # Went through all text that is part of an open bracket/close bracket
            # in other nodes
            # need to remove this text completely
            node.text = ""
        inside_openbrac_node = False
 
def unzip(source_filename, dest_dir):
    with ZipFile(source_filename) as zf:
        zf.extractall(dest_dir)
        
import tempfile,shutil

def _write_and_close_docx (xml_content, output_filename):
    """ Create a temp directory, expand the original docx zip.
        Write the modified xml to word/document.xml
        Zip it up as the new docx
    """

    tmp_dir = tempfile.mkdtemp()

    ZipFile.extractall(tmp_dir)

    with open(os.path.join(tmp_dir,'word/document.xml'), 'w') as f:
        xmlstr = etree.tostring (xml_content, pretty_print=True)
        f.write(xmlstr)

    # Get a list of all the files in the original docx zipfile
    filenames = ZipFile.namelist()
    # Now, create the new zip file and add all the filex into the archive
    zip_copy_filename = output_filename
    with ZipFile(zip_copy_filename, "w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmp_dir,filename), filename)

    # Clean up the temp dir
    shutil.rmtree(tmp_dir)

    
class OpenXML:
    def __init__(self):
        pass

import os 

curdir = os.getcwd()
filename = 'test.docx'#'blank.docx'#
filepath_in = os.path.join(curdir,filename)
filepath_out = os.path.join(curdir,'out.docx')

def main():
    xml =  get_word_xml(filepath_in)
    xml_tree = get_xml_tree(xml)
    #printxmltree(xml_tree)
    _join_tags(xml_tree)
    for node, txt in _itertext(xml_tree):
        print(txt)
    #_write_and_close_docx(filepath_out)
        
if __name__ == "__main__":
    main()

