# -*- coding: utf-8 -*-
"""
Created on Thu Mar 03 21:36:45 2016

@author: cduursma
"""

import docx2txt
from nltk.corpus import stopwords
#from nltk.stem.lancaster import LancasterStemmer
from nltk.stem.porter import PorterStemmer
import numpy as np
import csv

# this is ugly, shoudl be made flexible
MAX_KEYWORDS=1000
MY_DELIMITER=';'
MY_LANGUAGE='english'

#word_matrix=np.zeros((MAX_KEYWORDS), dtype='int32')
word_matrix=np.array([])
files_processed=[]
dictionary_header=[]
dictonary_full=False

puncList = [" ",".",";",":","!","?","/","\\",",","#","@","$","&",")","(","\"", "'"]
word_dict={}
stops = set(stopwords.words(MY_LANGUAGE))
st = PorterStemmer()
current_word_index=0


def get_word_index(w):
    global current_word_index
    global word_dict
    global dictionary_header
    global dictonary_full
    if w in word_dict:
        x=word_dict[w]
    else:
        if not dictonary_full:
            word_dict[w]=current_word_index
            x=current_word_index
            dictionary_header.append(w)
            current_word_index += 1
            if current_word_index > MAX_KEYWORDS:
                dictonary_full=True #Process conitnues but no new entries added
    return x
      
def vectorize_document(text):
    global word_matrix
    lines=text.splitlines()
    word_vector=np.zeros((MAX_KEYWORDS), dtype='int32')
    for line in lines:
        for w in line.split():
            clean_word=""
            for c in w:
                if not c in puncList:
                    if ord(c) < 129: # Only ASCII characters for now
                        clean_word=clean_word + c
            clean_word=clean_word.lower()
            if clean_word:
                if not clean_word.isdigit():
                    if clean_word not in stops:
                        clean_word=clean_word.lower()
                        stemmed_word=st.stem(clean_word)
                        word_index=get_word_index(stemmed_word) 
                        word_vector[word_index] += 1
    if word_matrix.any():
        word_matrix=np.vstack((word_matrix, word_vector))
    else: # processing first document
        word_matrix=word_vector
                        
def process_doxc_file(file_path):
    global files_processed
    text = docx2txt.process(file_path)
    vectorize_document(text)
    files_processed.append(file_path)
      
    
def write_header():
    global word_dict
    global word_matrix  
    nr_keys=len(dictionary_header)
    with open('matrix.csv', 'wb') as outcsv:
        writer = csv.writer(outcsv, delimiter=MY_DELIMITER)
        writer.writerow(dictionary_header)
    return nr_keys
    
def write_contents(nr_keys):
    global word_dict
    global word_matrix  
    print(nr_keys)
    reduced_matrix=word_matrix[0:len(files_processed), 0:nr_keys]
    with open('matrix.csv','a') as f_handle:
        np.savetxt(f_handle, reduced_matrix, fmt='%u', delimiter=MY_DELIMITER)   
  

def write_path_names():
    nr_files=len(files_processed)
    files_processed_vec = np.asarray(files_processed)
    files_processed_vec = files_processed_vec.reshape(nr_files,1);  
    with open('path_names.csv','wb') as f_handle:
        np.savetxt(f_handle, files_processed_vec, fmt='%s', delimiter=MY_DELIMITER)   
        
def get_files_processed():
    global files_processed
    return files_processed
  
