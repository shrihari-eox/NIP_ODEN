# -*- coding: utf-8 -*-
"""
Created on Sun Mar  6 21:20:26 2022

@author: Shreyas
"""

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

warnings.simplefilter(action="ignore")

def pdf_download():
    while len(os.listdir(r'Output')) == 0:
        print('waiting to download...')
        
    f = ''
    while f.split('.')[-1] != 'pdf':
        for file in os.listdir(r"Output"):
            f = file
        time.sleep(2)
        print('still waiting .....')
    print('----DOWNLOADED----')
    
def switch_to_right_frame():
    driver.switch_to.default_content()
    driver.switch_to.frame('RightFrame')

def switch_to_left_frame():
    driver.switch_to.default_content()
    driver.switch_to.frame('LeftFrame')


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
    merger.write(r"order_pdf\Combined_"+str(today)+".pdf")
    merger.close()

def error_mail(pNum):
    
    sender_email = "automation@eoxvantage.com"
    receiver_email = "ikhan@nipgroup.com,sjalal@nipgroup.com,pradeep@vantageagora.com,shriharim@eoxvantage.com"
    password = "Welcome2eox"
    message = MIMEMultipart("alternative")
    message["Subject"] = "NIP ODEN Error - "+str(pNum)
    message["From"] = sender_email
    message["To"] = receiver_email
    text = """
    Hi,

    The bot has found the error while running the Policy Number:"""+str(pNum)+""""

    Thanks and Regards,
    EOX Vantage

    """

    part1 = MIMEText(text, "plain")


    message.attach(part1)


        # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, [receiver_email], message.as_string()
        )


def completed_mail():
    sender_email = "automation@eoxvantage.com"
    receiver_email = "ikhan@nipgroup.com,sjalal@nipgroup.com,pradeep@vantageagora.com,shriharim@eoxvantage.com"
    password = "Welcome2eox"
    message = MIMEMultipart("alternative")
    message["Subject"] = "NIP ODEN Completed"
    message["From"] = sender_email
    message["To"] = receiver_email
    text = """
                Hi,
    
                The completed the processing the NIP ODEN .
    
                Thanks and Regards,
                EOX Vantage
    
                """
    
    part1 = MIMEText(text, "plain")
    
    
    message.attach(part1)
    
    
                    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
                        sender_email, [receiver_email], message.as_string()
                    ) 

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


# =============================================================================
#                           REQUIRED DETAILS
# =============================================================================
username = 'Pkumar'
password = 'VAdeveloper@1'
group = 'albiez'
df=pd.read_excel(r"C:\NIP ODEN\Input\Policy Data.xlsx",sheet_name="Sheet1")
all_details = pd.read_excel(r"C:\NIP ODEN\Input\AllDetails.xlsx","StateSpecificRulesState")



def NIP_ODEN_Run():
    driver.find_element_by_xpath('//*[@id="logonDialog"]/form/table/tbody/tr[1]/td/input').send_keys(username)
    driver.find_element_by_xpath('//*[@id="logonDialog"]/form/table/tbody/tr[2]/td/input').send_keys(password)
    driver.find_element_by_xpath('//*[@id="logonDialog"]/form/table/tbody/tr[3]/td/input').send_keys(group)
    driver.find_element_by_xpath('//*[@id="logonDialog"]/form/table/tbody/tr[4]/th/input[1]').click()
    
    driver.switch_to.frame('LeftFrame')
    WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menuDots"]/div[1]/form/table/tbody/tr[2]/td/a')))
    driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[2]/td/a').click()
    time.sleep(2)
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'1')))
    driver.find_element_by_name('1').click()
    WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//*[@id="menuDots"]/div[1]/form/table/tbody/tr[5]/td/a')))
    try:
        driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[5]/td/a').click()
    except:
        time.sleep(10)
        driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[5]/td/a').click()
    
    time.sleep(7)
    try:
        driver.switch_to.window(driver.window_handles[1])
    except:

        WebDriverWait(driver,10).until(EC.new_window_is_opened)
        try:
            driver.switch_to.window(driver.window_handles[1])
        except:
            driver.switch_to.window(driver.window_handles[1])
        

    
    driver.find_element_by_xpath('/html/body/form/table/tbody/tr/td[2]/input').send_keys(str(df["Policy Number"][i]))
    driver.find_element_by_xpath('/html/body/form/input[1]').click()
    
    result=driver.find_element_by_xpath('/html/body/div').text

    ''''Adding the New Policy'''

    if result=="No Records Found":
        print("No Policy")
        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[3]/td/a').click()
        driver.switch_to.default_content()
        driver.switch_to.frame("RightFrame")
        driver.find_element_by_name('sPolicyNumber').send_keys(str(df["Policy Number"][i]))
        driver.find_element_by_name('sNamedInsName1').clear()
        driver.find_element_by_name("sNamedInsName1").send_keys(str(df["Named Insured Line 1"][i]))
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        driver.find_element_by_partial_link_text("Continue").click()
    else:
        driver.find_element_by_xpath(
                '/html/body/table/tbody/tr[2]/td[1]/a/img').click()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(2)
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        while True:
            result1=driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[1]/td').text
            if "Commercial Policy" in result1:
                driver.find_element_by_partial_link_text("Edit").click()
                break    
    
    driver.switch_to.default_content()
    driver.switch_to.frame('RightFrame')
    state_code = str(df["State"][i]).strip()
    carrier_codeinput = str(df["NAIC Number"][i]).strip()
    # time.sleep(3)
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'sPolicyState')))
    selectstate_dropdwon = driver.find_element_by_xpath('/html/body/form/table[3]/tbody/tr[2]/td[2]/select')
    state_name = Select(selectstate_dropdwon)
    state_name.select_by_value(str(df["State"][i]).strip())
    policytype_dropdown = driver.find_element_by_xpath('/html/body/form/table[3]/tbody/tr[3]/td[2]/select')
    policy_type = Select(policytype_dropdown)
    policy_type.select_by_visible_text(str(df["Policy Title"][i]))
    driver.find_element_by_name('sNamedInsMSC').clear()
    name_line1 = driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[2]/td[2]/input').clear()
    driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[2]/td[2]/input').send_keys(str(df["Named Insured Line 1"][i]).strip())
    adress_line1 = driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[4]/td[2]/input').clear()
    driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[4]/td[2]/input').send_keys(str(df["Address Line 1"][i]).strip())
    city_name = driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[6]/td[2]/input').clear()
    driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[6]/td[2]/input').send_keys(str(df["City"][i]).strip())
    # time.sleep(2)
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'sState')))
    state1_dropdown = driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[7]/td[2]/select')
    state_1 = Select(state1_dropdown)
    state_1.select_by_value(str(df["State1"][i]).strip())
    driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[8]/td[2]/input').clear()
    if len(str(df["Zip Code"][i]).strip())<5:
        zip_code = driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[8]/td[2]/input').send_keys("0"+str(int(df["Zip Code"][i])).strip())
    else:
        zip_code = driver.find_element_by_xpath('/html/body/form/table[5]/tbody/tr[8]/td[2]/input').send_keys(str(int(df["Zip Code"][i])))
    
                                ###########################
    if int(df["NAIC Number"][i])==19518:
        driver.find_element_by_name("iNAICID").send_keys("Greenwich Insurance Company | 19518 | 0")
    elif int(df["NAIC Number"][i])==37885:
        driver.find_element_by_name("iNAICID").send_keys("XL Specialty insurance Co. | 37885 | 0")
    elif int(df["NAIC Number"][i])==24554:
        driver.find_element_by_name("iNAICID").send_keys("XL Insurance America, Inc. | 24554 | 0")
    elif int(df["NAIC Number"][i])==36940:
        driver.find_element_by_name("iNAICID").send_keys("Indian Harbor Insurance Co. | 36940 | 0") 
    
    
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'BranchName')))
    producernumber=driver.find_element_by_name("BranchName")
    nip=Select(driver.find_element_by_name("BranchName"))
    nip.select_by_index(3)
    for missing_data in range(1,10):
        time.sleep(3)
        pcode=driver.find_element_by_name("sProducerNbr").get_attribute('value')
        print(pcode)
        if pcode=="001":
            break
        else:
            driver.find_element_by_name("sProducerNbr").send_keys(Keys.CONTROL+'a')
            driver.find_element_by_name("sProducerNbr").send_keys(Keys.BACKSPACE)
            driver.find_element_by_name("sProducerNbr").send_keys('001',Keys.TAB)
            #driver.find_element_by_name("sProducerNbr").send_keys()
    
    driver.find_element_by_xpath('/html/body/form/table[13]/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/span/nobr/input').clear()
    inception_date = driver.find_element_by_xpath('/html/body/form/table[13]/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/span/nobr/input').send_keys(str (df["Inception Date"][i].strftime('%m/%d/%Y')).split(' ')[0])
    driver.find_element_by_xpath('/html/body/form/table[13]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/nobr/input').clear()
    exp_date = driver.find_element_by_xpath('/html/body/form/table[13]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/nobr/input').send_keys(str(df["Expiration Date"][i].strftime('%m/%d/%Y')).split(' ')[0])
    driver.switch_to.default_content()
    driver.switch_to.frame('LeftFrame')
    create_notice = driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[9]/td/a').click() 

    try:
        Alert(driver).accept()
    except:
        pass
    
    try:
        time.sleep(3)
        driver.find_element_by_partial_link_text("Continue").click()
    except:
        pass
                
    if df['State'][i].strip() not in ['VT','FL']:
        #Conditional Renewal selection
        time.sleep(3)
        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,'Conditional Renewal')))
        driver.find_element_by_partial_link_text('Conditional Renewal').click()
        try:
            Alert(driver).accept()
        except:
            pass
    
        # =============================================================================        
        #                  MAPPING FOR NEW JERSEY, WASHINGTON and ILLINIOS
        # =============================================================================   
        
        if df['State'][i].strip() in ['NJ','WA','IL']:
            print(df['State'][i].strip())
            switch_to_right_frame()
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table[2]/tbody')))
            if df['State'][i].strip() == 'IL':
                for j in range(len(all_details)):
                   if df["State"][i].strip()==all_details["State"][j].strip():
                       RC=all_details["Reason code selection"][j]
                       rowNum = j
                driver.find_element_by_name('sReasonID').send_keys(RC)
                
            switch_to_left_frame()
            driver.find_element_by_xpath('/html/body/div/div[1]/form/table/tbody/tr[6]/td/a').click()
            switch_to_right_frame()
            
            #map the written premium and date
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'p68_S')))
            driver.find_element_by_name('p68_S').send_keys(str('$')+ str("{:,}".format(int(df['Written premium'][i]))).split('.')[0]+' estimated. Premium is subject to rate change and underwriting')
            if df['State'][i].strip() in ['NJ','WA']:
                expDate = datetime.strptime(str(df['Expiration Date'][0]),'%Y-%m-%d %H:%M:%S')
                driver.find_element_by_name('p137_D').send_keys(str(expDate.month)+'/'+str(expDate.day)+'/'+str(expDate.year))
            
            switch_to_left_frame()
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[6]/td/a').click()
            
            #map terms and desc
            switch_to_right_frame()
            rowNum = all_details[all_details['State'].apply(lambda x: x.strip()) == df['State'][i].strip()].index[0]
            
                
        
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'mReasonDesc')))
            driver.find_element_by_name('mReasonDesc').send_keys(all_details['Terms and Description'][rowNum])
            switch_to_left_frame()
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[6]/td/a').click()
            
            
            #Calculate the max adv days
            switch_to_right_frame()
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table[2]/tbody/tr[6]/td[2]/select')))
            advDysNotice = driver.find_element_by_name('iAdvDays_N').get_attribute('value')
            if int(advDysNotice) < 1:
                print('Advance days notice is 0..')
                driver.find_element_by_name('iMailLeadDays_N').clear()
                Alert(driver).accept()
                driver.find_element_by_name('iMailLeadDays_N').click()
                driver.find_element_by_name('iMailLeadDays_N').send_keys(str(all_details['Mail Lead Time'][rowNum]))
                
                #click max adv days btn
                driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[4]/td[2]/div/input').click()
                Alert(driver).accept()        
                driver.find_element_by_name('iAdvDays_N').send_keys(str(all_details['Advance Days'][rowNum]))
            
            #click max adv days btn
            driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[4]/td[2]/div/input').click()
                
            #map mail and termination date
            try:
                Alert(driver).accept()
            except:
                pass
            mailType = driver.find_element_by_name('iMailTypeID')
            Select(mailType).select_by_visible_text(all_details['Mail Type'][rowNum])
            try:
                Alert(driver).accept() 
            except:
                pass
            effTimeOfNotice = driver.find_element_by_name('sAsOfPhrase')

            Select(effTimeOfNotice).select_by_visible_text(all_details['Effective Time of Notice'][rowNum])
            try:
                Alert(driver).accept() 
            except:
                pass
            switch_to_left_frame()
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[6]/td/a').click()
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="menuDots"]/div[1]/form/table/tbody/tr[7]/td/a')))
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[7]/td/a').click()
               
        
        
        
        
        
        # =============================================================================        
        #                       MAPPING FOR OTHER STATES FROMAT - CA
        # =============================================================================  
        else:
            print(df['State'][i].strip())
            switch_to_right_frame()
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table[2]/tbody')))
            
            for j in range(len(all_details)):
               if df["State"][i].strip()==all_details["State"][j].strip():
                   RC=all_details["Reason code selection"][j]
                   rowNum = j
            try:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table[2]/tbody/tr[2]/td[2]/select')))
                RCSelect = driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[2]/td[2]/select')
                Select(RCSelect).select_by_visible_text(str(RC).strip())
            except:
                pass
            
        # =============================================================================
        #                           FOR OKLAHOMA STATE (OK) and NEW YORK (NY)
        # =============================================================================
            if df['State'][i].strip() in ['OK','NY'] :
                switch_to_left_frame()
                driver.find_element_by_partial_link_text("Continue").click()
                switch_to_right_frame()
                if df['State'][i].strip() == 'OK':
                    try:
                        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'p131')))
                        driver.find_element_by_name('p131').send_keys(' ')
                    except:
                        print('In OK except')
                        pass
                else:
                    try:
                        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'p260')))
                        driver.find_element_by_name('p260').send_keys(' ')
                    except:
                        print('In NY except')
                        pass
             
            switch_to_left_frame()
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[6]/td/a').click()
            
            
        # =============================================================================
        #                           FOR ROHDE ISLAND (RI) STATE
        # =============================================================================
            if df['State'][i].strip() == 'RI':
                switch_to_right_frame()
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'p295069_S')))
                driver.find_element_by_name('p295069_S').send_keys(str('$')+str("{:,}".format(int(df['Written premium'][i]))).split('.')[0]+' estimated. Premium is subject to rate change and underwriting')
                switch_to_left_frame()
    
            
        # =============================================================================
        #                               FOR UTAH STATE (UT)
        # =============================================================================
            if df['State'][i].strip() == 'UT':
                switch_to_right_frame()
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'p68_S')))
                driver.find_element_by_name('p68_S').send_keys(str('$')+ str("{:,}".format(int(df['Written premium'][i]))).split('.')[0]+' estimated. Premium is subject to rate change and underwriting')
                driver.find_element_by_name('p295087').send_keys(' ')
                switch_to_left_frame()
    
    
            counter = 0
            while counter < 20:
                #Click on Contiue until the menu is Reason details
                try:
                    question=driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[1]/td').text
                except:
                    time.sleep(0.5)
                    counter += 1
                    question=driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[1]/td').text
                
                if question=="Reason Details":
                    break
                else:
                    switch_to_left_frame()
                    time.sleep(2)
                    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,'Continue')))
                    driver.find_element_by_partial_link_text("Continue").click()
            switch_to_right_frame()
        
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'mReasonDesc')))
            driver.find_element_by_name('mReasonDesc').send_keys(all_details['Terms and Description'][rowNum])
            switch_to_left_frame()
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[6]/td/a').click()
            
            #calculate max adv days
            switch_to_right_frame()
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'iMailTypeID')))
            
            advDysNotice = driver.find_element_by_name('iAdvDays_N').get_attribute('value')
            if int(advDysNotice) < 1 and df['State'][i].strip() not in ['OK','RI','VT']:
                print('Advance days notice is 0..\n')
                driver.find_element_by_name('iMailLeadDays_N').clear()
                try:
                    Alert(driver).accept()
                except:
                    pass
                driver.find_element_by_name('iMailLeadDays_N').click()
                driver.find_element_by_name('iMailLeadDays_N').send_keys(str(all_details['Mail Lead Time'][rowNum]))
                
                try:
                    Alert(driver).accept()
                except:
                    pass
                try:
                    Alert(driver).accept()
                except:
                    pass
                
                #click max adv days btn
                driver.find_element_by_name('MaxAdvanceDays').click()
                try:
                    Alert(driver).accept() 
                except:
                    pass
                driver.find_element_by_name('iAdvDays_N').clear()
                try:
                    Alert(driver).accept() 
                except:
                    pass
                driver.find_element_by_name('iAdvDays_N').send_keys(str(all_details['Advance Days'][rowNum]))
                try:
                    Alert(driver).accept()
                except:
                    pass
                
            #click max adv days btn
            driver.find_element_by_name('MaxAdvanceDays').click()
            
            #map mail and termination date
            try:
                Alert(driver).accept()
            except:
                pass
            mailType = driver.find_element_by_name('iMailTypeID')
            Select(mailType).select_by_visible_text(all_details['Mail Type'][rowNum])
            try:
                Alert(driver).accept() 
            except:
                pass

            effTimeOfNotice = driver.find_element_by_name('sAsOfPhrase')
            Select(effTimeOfNotice).select_by_visible_text(all_details['Effective Time of Notice'][rowNum])
            
            try:
                Alert(driver).accept() 
            except:
                pass
            
            switch_to_left_frame()
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[6]/td/a').click()
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="menuDots"]/div[1]/form/table/tbody/tr[7]/td/a')))
            driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[7]/td/a').click()
            
    elif df['State'][i].strip() == 'VT':
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'134')))
        driver.find_element_by_name('134').click()
        print(df['State'][i].strip())
        switch_to_right_frame()
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table[2]/tbody')))
        for j in range(len(all_details)):
            if df["State"][i].strip() == all_details["State"][j].strip():
                RC = all_details["Reason code selection"][j]
                rowNum = j
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table[2]/tbody/tr[2]/td[2]/select')))
        RCSelect = driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[2]/td[2]/select')
        Select(RCSelect).select_by_visible_text(RC)
        
        switch_to_left_frame()
        driver.find_element_by_name('10009').click()
        
        switch_to_right_frame()
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'p68_S')))
        driver.find_element_by_name('p68_S').send_keys(str('$')+ str("{:,}".format(int(df['Written premium'][i]))).split('.')[0]+' estimated. Premium is subject to rate change and underwriting')
        driver.find_element_by_name('p278').send_keys(str(all_details['Terms and Description'][rowNum]))
             
        switch_to_left_frame()
        driver.find_element_by_name('10075').click()
        
        switch_to_right_frame()
        driver.find_element_by_name('MaxAdvanceDays').click()
        mailType = driver.find_element_by_name('iMailTypeID')
        Select(mailType).select_by_visible_text(all_details['Mail Type'][rowNum])
        effTimeOfNotice = driver.find_element_by_name('sAsOfPhrase')
        Select(effTimeOfNotice).select_by_visible_text(all_details['Effective Time of Notice'][rowNum])
        
        switch_to_left_frame()
        driver.find_element_by_name('10081').click()
        
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'10087')))
        driver.find_element_by_name('10087').click()
    
    
    # =============================================================================
    #                       FOR FLORIDA STATE
    # =============================================================================
    elif df['State'][i].strip() == 'FL':
        switch_to_left_frame()
       
        try:
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'133')))
            driver.find_element_by_name('133').click()
        except:
            print("Error")
        try:
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'10009')))
            driver.find_element_by_name('10009').click()
        except:
            print("Driver")
        try:
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'10075')))
            driver.find_element_by_name('10075').click()
        except:
            print("Driver")
        
        for j in range(len(all_details)):
            if df["State"][i].strip() == all_details["State"][j].strip():
                RC = all_details["Reason code selection"][j]
                rowNum = j
        switch_to_right_frame()
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'mReasonDesc')))
        driver.find_element_by_name('mReasonDesc').send_keys(str(all_details['Terms and Description'][rowNum]))
        
        switch_to_left_frame()
        driver.find_element_by_name('10069').click()
        
        switch_to_right_frame()
        driver.find_element_by_name('iAdvDays_N').clear()
        try:
            Alert(driver).accept()
        except:
            print("Not there")
        driver.find_element_by_name('iAdvDays_N').send_keys(str(df['Days'][i]))
        mailType = driver.find_element_by_name('iMailTypeID')
        Select(mailType).select_by_visible_text(all_details['Mail Type'][rowNum])
        effTimeOfNotice = driver.find_element_by_name('sAsOfPhrase')
        Select(effTimeOfNotice).select_by_visible_text(all_details['Effective Time of Notice'][rowNum])
        
        switch_to_left_frame()
        driver.find_element_by_name('10081').click()
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,'10087')))
        driver.find_element_by_name('10087').click()
        
        
    pdf_download()
    driver.quit()
    
    for file in os.listdir(r"Output"):
        if file.endswith(".pdf"):
            shutil.move("Output/"+str(file), "ODEN Generated PDFs/"+str(df["Account"][i])+"_"+str(df["Policy Number"][i])+"_"+str(df["State"][i])+".pdf")
    df["Status"][i]="Generated" 



for i in range(len(df)):
    optns = Options()
    optns.add_experimental_option("useAutomationExtension", False)
    optns.add_experimental_option("excludeSwitches",["enable-automation"])
    prefs = {"credentials_enable_service": False,
     "profile.password_manager_enabled": False}
    optns.add_experimental_option("prefs", prefs)
    optns.binary_location = r"C:\Program Files\Google\chrome.exe"
    optns.add_argument(r"user-data-dir=C:\NIP ODEN\chromedata")
    chromedriver=r"C:\NIP ODEN\chromedriver.exe"
    driver = webdriver.Chrome(chromedriver, options=optns)
    driver.get("https://www.odenpt.com/ptlogin.asp")
    driver.maximize_window()
    
    try:
        NIP_ODEN_Run()
    except Exception as e:
        driver.quit()
        driver = webdriver.Chrome(chromedriver, options=optns)
        driver.get("https://www.odenpt.com/ptlogin.asp")        
        driver.maximize_window()
        try:
            NIP_ODEN_Run()
        except:
            #to note the exception type and line number
            exc_type, exc_obj, exc_tb = sys.exc_info()
            template = "An exception of type {0} occurred. Arguments:\n{1!r}"
            message = template.format(type(e).__name__, e.args)
            message = message + 'at line ----'+str(exc_tb.tb_lineno)
            print(message)
            df["Status"][i]="Not Generated - " +str(message)
            driver.quit()
            error_mail(df['Policy Number'][i])
        
    df.to_excel("output.xlsx",index=False)

# =============================================================================
#                       merge pdfs based on account id
# =============================================================================

merge_pdf_by_account_id()


# =============================================================================
#                       combine the pdfs based on time
# =============================================================================

combine_pdfs()
Conslidate_data()
completed_mail()
