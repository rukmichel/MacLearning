# -*- coding: utf-8 -*-
"""
Created on Thu May 26 12:37:22 2016

@author: mutabesham
"""
import win32com.client
import os
#from extract_text import write_header, write_contents, write_path_names, get_files_processed, vectorize_document, files_processed
class fileHandler:
    """"
    Common class for all type of document wrapper, f.e: wordWrapper, ExcelWarpper..
    """
    #all variables 
    #global file_paths  # List which will store all of the full filepaths.
    #global filePath
    #global type_app
    
    def __init__(self, path, typOfapp = "word.Application"):
         self.type_app = typOfapp
        
    def __get_list_paths__(self,path):
        file_paths = [] #instantiate the list of the full filepaths
        if os.path.isdir(path): # if the directory path not provided       
            # Walk the tree.
            for root, directories, files in os.walk(path):
                for filename in files:
                    # Join the two strings in order to form the full filepath.
                    filepath = os.path.join(root, filename)
                    file_paths.append(filepath)  # Add it to the list.
            return file_paths
            #if typOfapp == "" we can check if the application type is the type we have implemented        
        elif os.path.isfile(path):
            file_paths.append(path)
            return file_paths # this case it contain only one filepath
        else :
            raise fileHandler_Error('Not a valide path')
            
    def get_doc_properties(self,worddoc):
        try:
            csp2= worddoc.BuiltInDocumentProperties("Last Author").value
            print('Last author: %s' % csp2)
        except Exception as e:
            print ('\n\n', e)
        try:
            csp2= worddoc.BuiltInDocumentProperties("Language").value
            print('Title: %s' % csp2)
        except Exception as e:
            print ('\n\n', e)
        try:
            csp2= worddoc.BuiltInDocumentProperties("Number of Words").value
            print('Number of words: %s' % csp2)
        except Exception as e:
            print ('\n\n', e)

    def upgrade_Doc_ToDocx(self,pathFolder):
        """
        this function takes folder/file path and save all word document 
        in doc version to new version docx in the same forlder and keep the doc version too
        """
        try:
            wordApp = win32com.client.Dispatch(self.type_app) # instatiate the application
            wordApp.Visible = 0 # make wordapp work in background
            full_file_paths = self.__get_list_paths__(pathFolder) 
            for this_path in full_file_paths: 
                if this_path.endswith(('.doc')) and not this_path.startswith(('~')):
                    tempDoc = wordApp.Documents.Add(Template= this_path) # get exactly the copy of the file
                    tempDoc.SaveAs(this_path.replace('.doc', 'plaintext.txt'))
            wordApp.Close()
        except Exception as e:
            print(e.args)
        except fileHandler_Error as e:
            print(e.args)
        finally:
            
            return True
#    def get_file_ftp(self, destFolder)
            
import hashlib
import docx
class wordDocumentWrapper(fileHandler):
    
    
    def __init__(self, path, typOfapp = "word.Application"):
        super().__init__(path = path) # the same constractor as the parent class
    
    def __secureNameOnCV__():
        """
        This function replace all name with their equivalent sha1 value and save the name in CSV.
        """
        list_Names = []
        name_hashed = hashlib.sha1(nameInDoc)
        list_Names.append(nameInDoc)
    
    def splitFile(self, path , *args):
        """
        This function will user user the token to detect the end of english version of the CV and 
        and save all previous pages of the file. the the rest of the page will be saved as Dutch version.
        """
        token_word = '5 – Expert and capable as architect / advisor.' # this caracterize the end of the english CV.
        
        try:
            wordApp = win32com.client.Dispatch(self.type_app) # instatiate the application
            wordApp.Visible = 0 # make wordapp work in background
            if os.path.isfile(path):
                myDoc = wordApp.Documents.Add(path)
                doc = docx.Document(path)
                #text = myDoc.ActiveDocument.Sections[0].Headers[win32.constants.wdHeaderFooterPrimary].Range.Text
                #print(text)
                for para in doc.paragraphs:
                    for run in para:
                      print(run)
            wordApp.Close()
        except Exception as e:
            print(e.args)
        except fileHandler_Error as e:
            print(e.args)
        finally:
            return True

 ############################## code for client app goes here ######################################
        
#word_doc = wordDocumentWrapper('C:\\Users/mutabesham\\Documents')
#path = 'C:\\Users\\mutabesham\\Documents\\test.doc'
#listf = word_doc.upgrade_Doc_ToDocx(path)
#wordAp = win32com.client.Dispatch("word.Application")
#tempDoc = wordAp.Documents.Add(path)
## This tuple will determine the end of the english version CV.
#token_tuple = ('1 – Basic knowledge and limited experience.','2 – Average knowledge level with reasonable experience.','3 – Experienced.','4 – Very experienced and capable as coach.','5 – Expert and capable as architect / advisor.') 
##word_doc.splitFile(path, token_tuple)
##word_doc.get_doc_properties(tempDoc)
#
#help(docx.Document)
#

#
#import zipfile, lxml.etree
#
## open zipfile
#zf = zipfile.ZipFile('C:\\Users/mutabesham\\Documents\\test.docx')
# use lxml to parse the xml file we are interested in
#doc = lxml.etree.fromstring(zf.read('docProps/core.xml'))
# retrieve creator
#ns={'dc': 'http://purl.org/dc/elements/1.1/'}
#creator = doc.xpath('//dc:creator', namespaces=ns)[0].text


dir(wordApp.Documents.SaveAs())