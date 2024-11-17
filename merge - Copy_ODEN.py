from selenium.webdriver.common.keys import Keys
import shutil,os
from selenium import webdriver
from selenium.webdriver.support.ui import Select 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support  import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime
import time
from selenium.webdriver.common.alert import Alert
from PyPDF2 import PdfFileMerger
import stat
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib, ssl
import sys
import warnings
from pandas.core.common import SettingWithCopyWarning
import numpy as np
df=pd.read_excel(r"Input\Policy Data.xlsx",sheet_name="Sheet1")
def merge_pdf_by_account_id():
    lookup=df["Account"].values.tolist()
    list_of_unique_numbers = []
    
    lookup = set(lookup)
    
    for lookups in lookup:
        list_of_unique_numbers.append(lookups)
        
    for j in range(len(list_of_unique_numbers)):
        test=[]
        merger=PdfFileMerger()
    
        for i in os.listdir(r"ODEN Generated PDFs"):
            if list_of_unique_numbers[j]==str(i.split('_')[0]):
                test.append(r"ODEN Generated PDFs\/"+str(i))
        for page in test:
            merger.append(page)
        merger.write(r"Account\/"+str(list_of_unique_numbers[j])+".pdf")
        merger.close()

def combine_pdfs():
    root = os.listdir()
    path = r"ODEN Generated PDFs"
    filesSortedByDate = []
    sortedFilepath = []
    
    
    filepaths = [os.path.join(path, file) for file in os.listdir(path)]
    
    fileStatuses = [(os.stat(filepath), filepath) for filepath in filepaths]
    
    files = ((status[stat.ST_MTIME], filepath) for status, filepath in fileStatuses if stat.S_ISREG(status[stat.ST_MODE]))
    
    for modifiedTime, filepath in sorted(files):
        creation_date = time.ctime(modifiedTime)
        filename = os.path.basename(filepath)
        filesSortedByDate.append(creation_date + " " + filename)
        sortedFilepath.append(path+'\\'+filename)
    
    merger = PdfFileMerger()
    print(sortedFilepath)
    for pdf in sortedFilepath:
        merger.append(pdf)
    today = date.today()
    if "order_pdf" not in root:
        os.mkdir("order_pdf")
        print("order_pdf folder created")
    else:
        print("order_pdf folder exists")
    
    merger.write(r"order_pdf\Combined_"+str(today)+".pdf")
    merger.close()

merge_pdf_by_account_id()
combine_pdfs()