# -*- coding: utf-8 -*-
"""
Created on Wed Mar 09 21:05:49 2016

@author: cduursma
"""
# from http://mil-oss.org/resources/mil-oss-wg4-automating-microsoft-office-with-python.pdf 
import win32com.client, pythoncom
import os
from extract_text import write_header, write_contents, write_path_names, get_files_processed, vectorize_document, files_processed
print("test")
DOCUMENT_ROOT_PATH='C:\\Users\\mutabesham\\Documents'

pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

wordapp = win32com.client.Dispatch("Word.Application")

def get_filepaths(directory):
    """
    This function will generate the file names in a directory 
    tree by walking the tree either top-down or bottom-up. For each 
    directory in the tree rooted at directory top (including top itself), 
    it yields a 3-tuple (dirpath, dirnames, filenames).
    """
    file_paths = []  # List which will store all of the full filepaths.

    # Walk the tree.
    for root, directories, files in os.walk(directory):
        for filename in files:
            # Join the two strings in order to form the full filepath.
            filepath = os.path.join(root, filename)
            file_paths.append(filepath)  # Add it to the list.

    return file_paths  # Self-explanatory.

# Run the above function and store its results in a variable.   
full_file_paths = get_filepaths(DOCUMENT_ROOT_PATH)

for letter in full_file_paths:     # First Example
   print ('Current Letter :', letter)

def get_doc_properties(worddoc):
#    try:
#        csp= worddoc.CustomDocumentProperties().value
#        print('property is %s' % csp)
#    except Exception as e:
#        print ('\n\n', e)
    try:
        csp2= worddoc.BuiltInDocumentProperties("Last Author").value
        print('Last author: %s' % csp2)
    except Exception as e:
        print ('\n\n', e)
    try:
        csp2= worddoc.BuiltInDocumentProperties("Title").value
        print('Title: %s' % csp2)
    except Exception as e:
        print ('\n\n', e)
    try:
        csp2= worddoc.BuiltInDocumentProperties("Number of Words").value
        print('Number of words: %s' % csp2)
    except Exception as e:
        print ('\n\n', e)
    
        
#wordapp.Visible = True
# "C:\Users\cduursma\Documents\Python Scripts\doc2cluster\testfile.docx"
#worddoc = wordapp.Documents.Open('C:\Users\cduursma\Documents\Hello.docx', False, False, False)


extensions = ['.doc','.docx']

for this_path in full_file_paths: 
    if this_path.endswith(('.doc','.docx')):
        if not this_path.startswith(('~')):
            print ("TTTTTTTTT")
            print (this_path)
 #           worddoc.Visible = 0
            worddoc=wordapp.Documents.Open(this_path, False, False, False)
            get_doc_properties(worddoc)
            textdoc=worddoc.Content.Text   
            print(textdoc)
#            vectorize_document(textdoc)
            files_processed.append(this_path)
            worddoc.Close()
           
        
        
if len(get_files_processed()) > 0:
        nr_keys=write_header()
        write_contents(nr_keys)
        write_path_names()
  

wordapp.Quit() # Close the Word Application
