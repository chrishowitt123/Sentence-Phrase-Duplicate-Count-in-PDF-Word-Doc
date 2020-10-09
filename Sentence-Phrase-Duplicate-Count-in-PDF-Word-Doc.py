# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 14:43:49 2020

@author: chris
"""
import pdfplumber
import pandas as pd
from tkinter import filedialog
from iteration_utilities import deepflatten
from collections import Counter
import pandas as pd
import os

import docx2txt
import itertools
import re
from rapidfuzz import fuzz
from nltk.tokenize import PunktSentenceTokenizer

filetype = input("Type 'word' for a Word doc or 'pdf' for a PDF: ").lower()

if filetype == 'pdf':
    
    print( "\n")  
    print("Welcome to PDF_SimStrings!")
    print( "\n")  
    print("Please choose your file")
    print( "\n") 

    filename = filedialog.askopenfilename()
            
    fn = filename.split('.')[0] #seperate filename from extention for f-string output files
    
    n = 1
    allText = []
    
    
    
    with pdfplumber.open(filename) as pdf: #extract text from desired page
    
        while len(pdf.pages) >= n:
            page = pdf.pages[n-1]
            text = page.extract_text()
            #print( "\n")  #test to check tht all pages are being read
            #print('Page No. ' + str(n)) 
            #print( "\n") 
            #print(text)    
            allText.append(text)
            n = n +1
    
    allTextList = []
    
    for i in allText:
        s = i.split("\n")
        allTextList.append(s)
    
    allTextFlat = list(deepflatten(allTextList, depth=1))
    
    repeaters  = Counter(allTextFlat)
    
    results = pd.DataFrame.from_dict(repeaters, orient='index')
    results.columns = ['No. of instances']  
    results = results[results['No. of instances']>1]
    results.sort_values(by=['No. of instances'],ascending=False, inplace=True)
    pd.set_option("display.max_rows", None, "display.max_columns", None)
    print("Results!")
    print( "\n")
    print(results)
    
    
    os.path.dirname(os.path.abspath(filename))
    
    results.to_excel(filename + '_PDF_SimStrings.xlsx')
    
    print( "\n")
    print("Your PDF_SimStrings report was successfully created!")
    
    
else:
    print( "\n")  
    print("Welcome to DOCX_SimStrings!")
    print( "\n")  
    print("Please choose your file")
    print( "\n") 
    
    filename = filedialog.askopenfilename()
            
    fn = filename.split('.')[0] #seperate filename from extention for f-string output files
        
    text = docx2txt.process(filename)
    sent_tokenizer = PunktSentenceTokenizer(text)
    sents = sent_tokenizer.tokenize(text)
    
    allTextList = []
    
    for sent in sents:
        s = sent.split("\n")
        allTextList.append(s)
    
    allTextFlat = list(deepflatten(allTextList, depth=1))
    allTextFlat = [ x.replace('\t', '') for x in allTextFlat ]
    allTextFlat = [x for x in allTextFlat if x.strip()] 
    
    repeaters  = Counter(allTextFlat)
    
    results = pd.DataFrame.from_dict(repeaters, orient='index')
    results.columns = ['No. of instances']  
    results = results[results['No. of instances']>1]
    results.sort_values(by=['No. of instances'],ascending=False, inplace=True)
    pd.set_option("display.max_rows", None, "display.max_columns", None)
    print("Results!")
    print( "\n")
    print(results)
    
    
    os.path.dirname(os.path.abspath(filename))
    
    results.to_excel(filename + '_DOCX_SimStrings.xlsx')
    
    print( "\n")
    print("Your DOCX_SimStrings report was successfully created!")

    
