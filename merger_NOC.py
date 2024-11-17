import pyautogui
from selenium import webdriver
from check_weekday import get_day_value
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
import pandas as pd
from datetime import date 
from datetime import datetime
from datetime import timedelta
import time
import datetime,stat
import json
from selenium.webdriver.chrome.options import Options
import sys
import os
import warnings
import datetime
import shutil
from datetime import datetime
from pandas._libs.tslibs.timestamps import Timestamp
import os
import pandas as pd
import numpy as np
from PyPDF2 import PdfFileMerger
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import warnings
warnings.simplefilter(action="ignore")

def pdf_download():

    while len(os.listdir(r'PDF Rename')) == 0:
        time.sleep(0.2)
        print('waiting to download...')
        
    f = ''
    while f.split('.')[-1] != 'pdf':
        for file in os.listdir(r"Output"):
            f = file
        time.sleep(2)
        print('still waiting .....')
    print('----DOWNLOADED----')


def combine_pdfs():

    path = r"PDF Download"
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
    merger.write(r"order_pdf\Combined_"+str(today)+".pdf")
    merger.close()

def Conslidate_data():

    df=pd.read_excel(r"output.xlsx",sheet_name="Sheet1")
    dfGen = df.loc[df['Status'] == 'Generated']
    dfCon = pd.DataFrame(columns=['Full Name','Company','Address1','Address2','City','State','ZipCode','Zip4','Search1','Search2','Search3','Search4','Search5'])
    dfCon['Full Name'] = dfGen['Policy Number']
    dfCon['Company'] = dfGen['Named Insured Line 1']
    dfCon['Address1'] = dfGen['Address Line 1']
    dfCon['City'] = dfGen['City']
    dfCon['State'] = dfGen['State']
    dfCon['ZipCode'] = dfGen['Zip Code']
    dfCon.replace(np.nan, '', regex=True)
    dfCon.to_csv(r'Consolidated.csv',index=False)

df=pd.read_excel(r"C:\NIP NOC\NOC Policy Data.xlsx",sheet_name="Sheet1")

def merge_pdf_by_account_id():
    
    lookup= df["Account"].values.tolist()
    list_of_unique_numbers = []
    
    lookup = set(lookup)
    
    for lookups in lookup:
        list_of_unique_numbers.append(lookups)
        
    for j in range(len(list_of_unique_numbers)):
        test=[]
        merger=PdfFileMerger()
    
        for i in os.listdir(r"C:\NIP NOC\PDF Download"):
            if list_of_unique_numbers[j] in str(i.split('_')[1]):
                test.append(r"C:\NIP NOC\PDF Download\/"+str(i))
        for page in test:
            merger.append(page)
        merger.write(r"C:\NIP NOC\PDF Account\/"+str(list_of_unique_numbers[j])+".pdf")
        merger.close()

merge_pdf_by_account_id()
print("pdfs merged")