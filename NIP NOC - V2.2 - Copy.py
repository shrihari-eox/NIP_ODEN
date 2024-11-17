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

def merge_pdf_by_account_id():
    
    lookup=df["Account"].values.tolist()
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

def error_mail(pNum):
    
    sender_email = "automation@eoxvantage.com"
    receiver_email = "ikhan@nipgroup.com,sjalal@nipgroup.com,pradeep@vantageagora.com,shriharim@eoxvantage.com"
    password = "Welcome2eox"
    message = MIMEMultipart("alternative")
    message["Subject"] = "NIP NOC Error - "+str(pNum)
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


df=pd.read_excel(r"C:\NIP NOC\NOC Policy Data.xlsx",sheet_name="Sheet1")
    
    
#df=pd.read_excel(r"C:\NIP NOC\Copy of NOC Policy Data N.V.xlsx",sheet_name="Sheet1")
df1=pd.read_excel(r"C:\NIP NOC\AllDetails.xlsx",sheet_name="StateSpecificRulesState")



for i in range(len(df["State"])):

    print(len(df["State"]) - i , 'policies left')
    # try:
    ''''-----------------DRIVER----------------------'''
    optns = Options()
    optns.add_experimental_option("useAutomationExtension", False)
    optns.add_experimental_option("excludeSwitches",["enable-automation"])
    prefs = {"credentials_enable_service": False,
    "profile.password_manager_enabled": False}
    optns.add_experimental_option("prefs", prefs)
    optns.binary_location = r"C:\Program Files\Google\chrome.exe"
    optns.add_argument(r"user-data-dir=C:\NIP NOC\chromedata")
    chromedriver=r"C:\NIP NOC\chromedriver.exe"
    
    driver = webdriver.Chrome(chromedriver, options=optns)
    driver.get("https://www.odenpt.com/ptlogin.asp")
    driver.maximize_window()

    '''------------------CREDENTIALS--------------------'''

    # driver.find_element_by_xpath(
    #     '//*[@id="logonDialog"]/form/table/tbody/tr[1]/td/input').send_keys('Program VA User')
    # driver.find_element_by_xpath(
    #     '//*[@id="logonDialog"]/form/table/tbody/tr[2]/td/input').send_keys('VA@NIPgroup1')
    # driver.find_element_by_xpath(
    #     '//*[@id="logonDialog"]/form/table/tbody/tr[3]/td/input').send_keys('albiez')
    
    driver.find_element_by_xpath(
            '//*[@id="logonDialog"]/form/table/tbody/tr[1]/td/input').send_keys('Pkumar')
    driver.find_element_by_xpath(
            '//*[@id="logonDialog"]/form/table/tbody/tr[2]/td/input').send_keys('VAdeveloper@1')
    driver.find_element_by_xpath(
            '//*[@id="logonDialog"]/form/table/tbody/tr[3]/td/input').send_keys('albiez')
    

   
    driver.find_element_by_xpath(
        '//*[@id="logonDialog"]/form/table/tbody/tr[4]/th/input[1]').click()
    driver.switch_to.frame("LeftFrame")

    time.sleep(5)

    driver.find_element_by_xpath(
        '//*[@id="menuDots"]/div[1]/form/table/tbody/tr[2]/td/a').click()
    time.sleep(0.8)
    driver.find_element_by_xpath(
        '/html/body/div/div[1]/form/table/tbody/tr[2]/td/a').click()
    time.sleep(0.8)
    driver.find_element_by_xpath(
        '//*[@id="menuDots"]/div[1]/form/table/tbody/tr[5]/td/a').click()

    '''-------------POLICY NUMBER------------------'''

    WebDriverWait(driver, 200).until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[1])
    WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr/td[2]/input')))
    driver.find_element_by_xpath(
        '/html/body/form/table/tbody/tr/td[2]/input').send_keys(str(df["Policy Number"][i]))
    driver.find_element_by_xpath('/html/body/form/input[1]').click()
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div')))
    result=driver.find_element_by_xpath('/html/body/div').text

    '''------------------POLICY FOUND OR NOT FOUND--------------------------'''

    if result=="No Records Found":
        print("No Policy")
        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[3]/td/a').click()
        driver.switch_to.default_content()
        driver.switch_to.frame("RightFrame")
        driver.find_element_by_name('sPolicyNumber').send_keys(str(df["Policy Number"][i]))
        driver.find_element_by_name("sNamedInsName1").send_keys(str(df["Named Insured Line 1"][i]))
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        driver.find_element_by_partial_link_text("Continue").click()
    else:
        driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[2]/td[1]/a/img').click()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(5)
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        driver.find_element_by_partial_link_text("Edit").click()
    driver.switch_to.default_content()
    driver.switch_to.frame("RightFrame")

    '''---------------STATE----------------'''


    try:
        State=Select(driver.find_element_by_name("sPolicyState"))
        State.select_by_value(str(df["State"][i]).strip())
    except:
        driver.find_element_by_name("sPolicyState").send_keys(str(df["State"][i]).strip())

    test=Select(driver.find_element_by_name("iUserPolicyTypeID"))
    test.select_by_visible_text(str(df["Policy Title"][i]))

    driver.find_element_by_name("sNamedInsMSC").clear()
    driver.find_element_by_name('sNamedInsMSC').send_keys("Certified Mail")
    
    '''                         Named Insured Line 1                '''

    driver.find_element_by_name("sNamedInsName1").clear()
    driver.find_element_by_name("sNamedInsName1").send_keys(str(df["Named Insured Line 1"][i]))

    ''''--------------ADDRESS LINE 1-------------'''
    driver.find_element_by_name("sAddr1").clear()
    driver.find_element_by_name("sAddr1").send_keys(str(df["Address Line 1"][i]))

    '''----------------CITY--------------------'''
    driver.find_element_by_name("sCity").clear()
    driver.find_element_by_name("sCity").send_keys(str(df["City"][i]))

    driver.find_element_by_name("sState").send_keys(str(df["State1"][i]).strip())
    
    driver.find_element_by_name("sZip").clear()
    if len(str(df["Zip Code"][i]).strip())<5:
        driver.find_element_by_name("sZip").send_keys("0"+str(df["Zip Code"][i]))
    else:
        driver.find_element_by_name("sZip").send_keys(str(df["Zip Code"][i]))
    
    
    '''---------------NAIC CODE-------------'''
    
    if str(df["NAIC Number"][i])=="19518":
        driver.find_element_by_name("iNAICID").send_keys("Greenwich Insurance Company | 19518 | 0")
    elif str(df["NAIC Number"][i])=="37885":
        driver.find_element_by_name("iNAICID").send_keys("XL Specialty insurance Co. | 37885 | 0")
    elif str(df["NAIC Number"][i])=="24554":
        driver.find_element_by_name("iNAICID").send_keys("XL Insurance America, Inc. | 24554 | 0")
    elif str(df["NAIC Number"][i])=="36940":
        driver.find_element_by_name("iNAICID").send_keys("Indian Harbor Insurance Co. | 36940 | 0")    
    
    '''-------------BRANCH NAME------------'''
    nip=Select(driver.find_element_by_name("BranchName"))
    nip.select_by_index(3)
    producernumber=driver.find_element_by_name("BranchName")


    '''--------------PRODUCER NUMBER---------'''

    for missing_data in range(1,5):
        time.sleep(1)
        pcode=driver.find_element_by_name("sProducerNbr").get_attribute('value')
        print(pcode)
        if pcode=="001":
            break
        else:  
            driver.find_element_by_name("sProducerNbr").send_keys(Keys.CONTROL + "a")
            driver.find_element_by_name("sProducerNbr").send_keys(Keys.DELETE)
            time.sleep(2)
            print("should be clear")
            driver.find_element_by_name("sProducerNbr").send_keys('001',Keys.TAB)
            
            
        
        
    '''---------------INCEPTION DATE-----------'''
    driver.find_element_by_name("dtInception").clear()
    
    inception=str(df["Inception Date"][i]).split(" ")[0].split("-")
    inception_date=str(inception[1])+"/"+str(inception[2])+"/"+str(inception[0])
    
    driver.find_element_by_name("dtInception").send_keys(inception_date)

    '''---------------EXPIRATION DATE-----------'''

    driver.find_element_by_name("dtExpiration").clear()
    exception=str(df["Expiration Date"][i]).split(" ")[0].split("-")
    exception_date=str(exception[1])+"/"+str(exception[2])+"/"+str(exception[0])
    driver.find_element_by_name("dtExpiration").send_keys(exception_date)


    '''------TOTAL $-------------------------'''
    if str(df["State"][i])=="NJ":
        try:
            driver.find_element_by_name("curPremiumDue").clear()
        except:
            print("Clear")
            
        
        try:
            driver.find_element_by_name("curPremiumDue").send_keys("{:,}".format(int(str(df["Total $"][i]))))
        except:
            driver.find_element_by_name("curPremiumDue").send_keys(str(df["Total $"][i]).split('.')[0])

        '''--------------DUE DATE-------------------'''
        desc_test1=str(df["AgeDate"][i]).split(" ")[0].split("-")
        desc_test1=str(desc_test1[1])+"/"+str(desc_test1[2])+"/"+str(desc_test1[0])
        
        try:
            driver.find_element_by_name("dtPremium").clear()
        except:
            print("Clear")
        driver.find_element_by_name("dtPremium").send_keys(desc_test1)


    driver.switch_to.default_content()
    
    driver.switch_to.frame("LeftFrame")

    '''-------------------CREATE NOTICE----------'''
    driver.find_element_by_link_text('Create Notice').click()

                
    '''-------------TEXAS CONDITION----------'''
    while True:
        try:
            driver.find_element_by_partial_link_text("Continue").click()
        except:
            print("Texas")
        try:
            driver.find_element_by_partial_link_text("Cancellation").click()
            break
        except:
            print("Click Cancellation again")
            
    
    driver.switch_to.default_content()
    driver.switch_to.frame("RightFrame")

    '''-----------------REASON CODE------------'''
    time.sleep(3)
    try:
        select1 = driver.find_element_by_name(
                "sReasonID").send_keys("NONPAYMENT OF PREMIUM")
    except:
        print("Select")
    driver.switch_to.default_content()
    driver.switch_to.frame("LeftFrame")
    driver.find_element_by_partial_link_text("Continue").click()
    time.sleep(2)


    '''----------------PREMIUM------------------'''

    driver.switch_to.default_content()
    driver.switch_to.frame("RightFrame")
    k=j=1
    try:
        try:
            driver.find_element_by_name('curPremiumDue').send_keys("{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
        except:
            driver.find_element_by_name('curPremiumDue').send_keys((str(df["Total $"][i]).split(".")[0]))
        k=0
    except:
        print("Current Premium")


    try:
        duedate=str(df["Due Date"][0]).split(" ")[0].split("-")
        duedate_final=str(duedate[1])+"/"+str(duedate[2])+"/"+str(duedate[0])
        driver.find_element_by_name('dtPremium').send_keys(duedate_final)
        j=0
    except:
        print("Current Premium")



    '''---------------NEWYORK CONDITION----------------------------'''

    if k==0 or j==0:
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
        driver.find_element_by_partial_link_text("Continue").click()
            
        driver.switch_to.default_content()
        driver.switch_to.frame("RightFrame")
    try:
        if df["State"] == 'NY':
            driver.find_element_by_name('p295073_S').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))
        else:
            driver.find_element_by_name('p295073_S').send_keys("{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
    except:
        try:
            driver.find_element_by_name('p295073_S').send_keys(str(df["Total $"][i]).split(".")[0])
        except:
            print("Current Premium")
        
    try:
        driver.find_element_by_name('p272').send_keys("800-446-7647 \nNIP Group Inc \n900 Route 9 North \nWoodbridge NJ 07095")
    except:
        print("Address Phone")
    
    '''---------------MASSACHUSETTS CONDITION----------------------------'''
    
    if str(df["State"][i]) == 'MA':
        time.sleep(2)
        driver.find_element_by_name('p298_S').click()
        print('clicked---?')
        driver.find_element_by_name('p298_S').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))
        driver.switch_to.default_content()
        driver.switch_to.frame("LeftFrame")
    else:
        print('did not execute')
        
    driver.switch_to.default_content()
    driver.switch_to.frame("LeftFrame")

    # if k==0 or j==0:
    #     driver.switch_to.default_content()
    #     driver.switch_to.frame("LeftFrame")
    #     driver.find_element_by_partial_link_text("Continue").click()
            
    #     driver.switch_to.default_content()
    #     driver.switch_to.frame("RightFrame")
    # try:
    #     if df["State"] == 'MA':
    #         driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[2]/td').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))

    #         #driver.find_element_by_name('p298_S').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))
    #     else:
    #         driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[2]/td').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))
    #         #driver.find_element_by_name('p298_S').send_keys("{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
    # except:
    #     try:
    #         driver.find_element_by_xpath('/html/body/form/table[2]/tbody/tr[2]/td').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))
    #         #driver.find_element_by_name('p298_S').send_keys(str(df["Total $"][i]).split(".")[0])
    #     except:
    #         print("Current Premium MA ")
        
    # try:
    #     driver.find_element_by_name('p272').send_keys("800-446-7647 \nNIP Group Inc \n900 Route 9 North \nWoodbridge NJ 07095")
    # except:
    #     print("Address Phone")
    # driver.switch_to.default_content()
    # driver.switch_to.frame("LeftFrame")
    

    '''----------------CLICKING CONTINUE------------------'''

    while True:
        
        time.sleep(1)
        question=driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[1]/td').text
        print(question)
        if question=="Reason Details" or question== "Mail & Termination Date Information":
            break
        else:
            driver.switch_to.default_content()
            driver.switch_to.frame("LeftFrame")
            driver.find_element_by_partial_link_text("Continue").click()
            print('clicked continue')
    driver.switch_to.default_content()
    driver.switch_to.frame("RightFrame")
    
    time.sleep(5)
    


    '''----------------------DESCRIPTION-------------------------'''
    checkDate = datetime.strptime('05-15-2022','%m-%d-%Y')
    cancelDate = datetime.today()+timedelta(days=25)
    cancelDate = cancelDate.strftime('%m/%d/%Y')
    reasonDetails = '''due to non-payment of premium of $'''+"{:,}".format(int(df['Total $'][i]))+''' which was due on '''+str(df['AgeDate'][i].strftime('%m/%d/%Y'))+''' \n

'''
    for desc in range(len(df1)):
        if str(df["State"][i]) in str(df1["State"][desc]) or str(df1["State"][desc]) in str(df["State"][i]):
            if df['Inception Date'][i] < checkDate:
                date_desc=str(date.today()+ timedelta(days=25)).split("-")
                date_desc=str(date_desc[1])+"/"+str(date_desc[2])+"/"+str(date_desc[0])
                descr=str(df1["Terms and Description"][desc]).replace("05/07/2020", date_desc)
                desc_test=str(df["AgeDate"][i]).split(" ")[0].split("-")
                desc_test=str(desc_test[1])+"/"+str(desc_test[2])+"/"+str(desc_test[0])
                expiration=descr.replace("02/01/2020", desc_test)
                try:
                    premium=expiration.replace("500","{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
                except:
                    premium=expiration.replace("500",(str(df["Total $"][i]).split(".")[0]))
                driver.find_element_by_name("mReasonDesc").send_keys(premium)
                break
            elif df['Inception Date'][i] >= checkDate:
                if str(df["State"][i]) != 'MA':
                    driver.find_element_by_name("mReasonDesc").send_keys(reasonDetails)

    driver.switch_to.default_content()
    driver.switch_to.frame("LeftFrame")
    if str(df["State"][i]) != 'MA':
        driver.find_element_by_partial_link_text("Continue").click()
    time.sleep(5)
    driver.switch_to.default_content()
    driver.switch_to.frame("RightFrame")

    print("finding ADN")
    
    

    ADN = driver.find_element_by_name('iAdvDays_N').get_attribute("value")
    print("ADN Found")
    print("ADN value extracted: " + ADN)
    if int(ADN) == 0:
        print("if loop entered") 
        ADN = 10
        print("ADN IS : " + str(ADN))
        
        driver.find_element_by_name('dtMail_D').send_keys(Keys.TAB, Keys.TAB, str(ADN))
        time.sleep(3)
        
    else: 
        print("Unchanged ADN value is: "+ str(ADN))
    '''----------------------MAIL DATE-----------'''
    
    try:
        driver.find_element_by_name('dtMail_D').clear()
        date_t=str(date.today()+ timedelta(days=1)).split("-")
        date_final=str(date_t[1])+"/"+str(date_t[2])+"/"+str(date_t[0])
        driver.find_element_by_name('dtMail_D').send_keys(date_final)
        
    except:
        date_t=str(date.today()+ timedelta(days=1)).split("-")
        date_final=str(date_t[1])+"/"+str(date_t[2])+"/"+str(date_t[0])
        try:
            MLTE= driver.find_element_by_name('iMailLeadDays_N')
            MLTVAL = MLTE.get_attribute("value")
            print(MLTVAL)
            if int(MLTVAL) == 0:
                MLTVAL = int(5)
                driver.find_element_by_name('dtMail_D').send_keys(date_final,Keys.TAB,Keys.ENTER)
                
                #sdriver.find_element_by_name('iMailLeadDays_N').send_keys(MLTVAL)
                #print("enter+TAB clicked")
                time.sleep(2)
                driver.find_element_by_name('iMailLeadDays_N').send_keys((MLTVAL), Keys.TAB)
                #time.sleep(2)
                #driver.find_element_by_name('iMailLeadDays_N').send_keys(Keys.ENTER)
            else:
                driver.find_element_by_name('dtMail_D').send_keys(date_final)
        except:
            print("No Mailing Date")
            driver.find_element_by_name('dtMail_D').send_keys(Keys.TAB,Keys.ENTER,MLTVAL)
        
######################################################################################## CHANGE ################################################################################################################
    


    
    
    #if df["State"][i]=="NY":
    #    try:
    #        driver.find_element_by_name('iMailLeadDays_N').clear()
    #        driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])-5))
    #    except:
    #        try:
    #            driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])-5))
    #        except:
    #            try:
    #                driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])-5))
    #            except:
    #                driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])-5))
    #else:
    #    try:
    #        driver.find_element_by_name('iMailLeadDays_N').clear()
    #        driver.find_element_by_name('iMailLeadDays_N').send_keys(str(df["Days"][i]))
    #    except:
    #        try:
    #            driver.find_element_by_name('iMailLeadDays_N').send_keys(str(df["Days"][i]))
    #        except:
    #            try:
    #                driver.find_element_by_name('iMailLeadDays_N').send_keys(str(df["Days"][i]))
    #            except:
    #                driver.find_element_by_name('iMailLeadDays_N').send_keys(str(df["Days"][i])) 
                    
  ###################################################################################### END OF CHANGE ##################################################################################################################              
    print("getting effective date....")
    #driver.find_element_by_name('dtMail_D').send_keys(Keys.TAB, Keys.ENTER)

    # print("presssed tab")
    # cancellation_date = driver.find_element_by_name('dtEffective_D').get_attribute("value")
    # print(cancellation_date)
    # print(type(date_final))
    # print(date_final)

    

    print(MLTVAL)

    strdate = date_final
    datetest = datetime.strptime(strdate, "%m/%d/%Y")
    dta = int(ADN) + int(MLTVAL) + 1 

    can_date = datetest + timedelta(days= dta)
    print(dta)
    print(can_date)
    can_date = can_date.strftime("%Y %m %d")
    can_date = datetime.strptime(can_date, "%Y %m %d")
    checkWeekVal = get_day_value(can_date)
    if checkWeekVal == 1:
        MLTVAL = str(int(MLTVAL) + 1)
        driver.find_element_by_name('dtMail_D').click()
        time.sleep(5)
        driver.find_element_by_name('dtMail_D').send_keys(Keys.TAB,Keys.ENTER,str(MLTVAL))
        
    elif checkWeekVal == 2:
        MLTVAL = str(int(MLTVAL) + 2)
        
        driver.find_element_by_name('dtMail_D').click()
        try:

            Alert(driver).accept()
        except:
            pass
        
        driver.find_element_by_name('dtMail_D').send_keys(Keys.TAB,str(MLTVAL))
        



    
    try:
        Alert(driver).accept()

    except:
        
        print("effective date")
        # time.sleep(100)

    for desc in range(len(df1)):
        if str(df["State"][i]) in str(df1["State"][desc]):
            try:
                driver.find_element_by_name('iAdvDays_N').clear()
                time.sleep(3)
                Alert(driver).accept()
                driver.find_element_by_name('iAdvDays_N').send_keys(str(df1["Advance Days"][desc]))
            except:
                try:
                    driver.find_element_by_name('iAdvDays_N').clear()
                    time.sleep(2)
                    Alert(driver).accept()
                    driver.find_element_by_name('iAdvDays_N').send_keys(str(df1["Advance Days"][desc]))
                except:
                    try:
                        driver.find_element_by_name('iAdvDays_N').clear()
                        time.sleep(2)
                        Alert(driver).accept()
                        driver.find_element_by_name('iAdvDays_N').send_keys(str(df1["Advance Days"][desc]))
                    except:
                        try:
                            driver.find_element_by_name('iAdvDays_N').clear()
                            time.sleep(2)
                            Alert(driver).accept()
                            driver.find_element_by_name('iAdvDays_N').send_keys(str(df1["Advance Days"][desc]))
                        except:
                            driver.find_element_by_name('iAdvDays_N').clear()
                            time.sleep(2)
                            try:
                                Alert(driver).accept()
                            except:
                                pass
                            driver.find_element_by_name('iAdvDays_N').send_keys(str(df1["Advance Days"][desc]))
                            
    
    try:
        Alert(driver).accept()
        driver.switch_to_alert().accept()
    except:
        print("certi2")
        time.sleep(5)
    try:
        
        driver.find_element_by_name('iMailTypeID').clear()
        driver.find_element_by_name('iMailTypeID').send_keys("Certified")
        
    except:
        try:
            driver.find_element_by_name('iMailTypeID').send_keys("Certified")
            
        except:
            driver.find_element_by_name('iMailTypeID').send_keys("Certified")
            

    try:
        Alert(driver).accept()
        driver.switch_to_alert().accept()
    except:
        print("Pop Up")
    try:
        
        driver.find_element_by_name('sAsOfPhrase').send_keys("12:01 A.M. Local Time")
        
    except:
        
        driver.find_element_by_name('sAsOfPhrase').send_keys("12:01 A.M. Local  Time")
    driver.switch_to.default_content()
    driver.switch_to.frame("LeftFrame")
    driver.find_element_by_partial_link_text("Continue").click()
    
    time.sleep(3)
    driver.find_element_by_partial_link_text("Generate").click()
    WebDriverWait(driver, 50).until(EC.number_of_windows_to_be(2))
    time.sleep(15)
    

    driver.quit()
    print(df["State"][i])

    '''------------RENAME FILE-----------'''    
    #################################################################################################### REVIEW #############################################################################################################################
    os.chdir("C:\\NIP NOC\\")
    for file in os.listdir(r"C:\NIP NOC\PDF Rename"):
        if file.endswith(".pdf"):
            src_path=str("PDF Rename\\" + str(file))
            dest_path= str("PDF Download\\" + str(df["Policy Number"][i]) +"_"+ str(df["Account"][i]) +"_"+ str(df["State"][i]) +".pdf")
            shutil.move(src_path,dest_path)
    df["Status"][i]="Generated" 
    # except:
    #     try:
    #         driver.quit()
    #     except:
    #         print("Browser already closed")
    #     try:
    #         ''''-----------------DRIVER----------------------'''
    #         optns = Options()
    #         optns.add_experimental_option("useAutomationExtension", False)
    #         optns.add_experimental_option("excludeSwitches",["enable-automation"])
    #         prefs = {"credentials_enable_service": False,
    #         "profile.password_manager_enabled": False}
    #         optns.add_experimental_option("prefs", prefs)
    #         optns.binary_location = r"C:\Program Files\Google\chrome.exe"
    #         optns.add_argument(r"user-data-dir=C:\NIP NOC\chromedata")
    #         chromedriver=r"chromedriver.exe"
    #         driver = webdriver.Chrome(chromedriver, options=optns)
    #         driver.get("https://www.odenpt.com/ptlogin.asp")
    #         driver.maximize_window()

    #         '''------------------CREDENTIALS--------------------'''

    #         driver.find_element_by_xpath(
    #             '//*[@id="logonDialog"]/form/table/tbody/tr[1]/td/input').send_keys('Program VA User')
    #         driver.find_element_by_xpath(
    #             '//*[@id="logonDialog"]/form/table/tbody/tr[2]/td/input').send_keys('VA@NIPgroup1')
    #         driver.find_element_by_xpath(
    #             '//*[@id="logonDialog"]/form/table/tbody/tr[3]/td/input').send_keys('Albiez')
            
    #         driver.find_element_by_xpath(
    #             '//*[@id="logonDialog"]/form/table/tbody/tr[4]/th/input[1]').click()
    #         driver.switch_to.frame("LeftFrame")

    #         time.sleep(3)

    #         driver.find_element_by_xpath(
    #             '//*[@id="menuDots"]/div[1]/form/table/tbody/tr[2]/td/a').click()
            
    #         driver.find_element_by_xpath(
    #             '//*[@id="menuDots"]/div[1]/form/table/tbody/tr[2]/td/a').click()
    #         driver.find_element_by_xpath(
    #             '//*[@id="menuDots"]/div[1]/form/table/tbody/tr[5]/td/a').click()

    #         '''-------------POLICY NUMBER------------------'''

    #         WebDriverWait(driver, 50).until(EC.number_of_windows_to_be(2))
    #         driver.switch_to.window(driver.window_handles[1])
    #         WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr/td[2]/input')))
    #         driver.find_element_by_xpath(
    #             '/html/body/form/table/tbody/tr/td[2]/input').send_keys(str(df["Policy Number"][i]))
    #         driver.find_element_by_xpath('/html/body/form/input[1]').click()
    #         WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div')))
    #         result=driver.find_element_by_xpath('/html/body/div').text

    #         '''------------------POLICY FOUND OR NOT FOUND--------------------------'''

    #         if result=="No Records Found":
    #             print("No Policy")
    #             driver.switch_to.window(driver.window_handles[0])
    #             driver.switch_to.default_content()
    #             driver.switch_to.frame("LeftFrame")
    #             driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[3]/td/a').click()
    #             driver.switch_to.default_content()
    #             driver.switch_to.frame("RightFrame")
    #             driver.find_element_by_name('sPolicyNumber').send_keys(str(df["Policy Number"][i]))
    #             driver.find_element_by_name("sNamedInsName1").send_keys(str(df["Named Insured Line 1"][i]))
    #             driver.switch_to.default_content()
    #             driver.switch_to.frame("LeftFrame")
    #             driver.find_element_by_partial_link_text("Continue").click()
    #         else:
    #             driver.find_element_by_xpath(
    #                 '/html/body/table/tbody/tr[2]/td[1]/a/img').click()
    #             driver.switch_to.window(driver.window_handles[0])
    #             time.sleep(5)
    #             driver.switch_to.default_content()
    #             driver.switch_to.frame("LeftFrame")
    #             driver.find_element_by_partial_link_text("Edit").click()
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("RightFrame")

    #         '''---------------STATE----------------'''


    #         try:
    #             State=Select(driver.find_element_by_name("sPolicyState"))
    #             State.select_by_value(str(df["State"][i]).strip())
    #         except:
    #             driver.find_element_by_name("sPolicyState").send_keys(str(df["State"][i]).strip())

    #         test=Select(driver.find_element_by_name("iUserPolicyTypeID"))
    #         test.select_by_visible_text(str(df["Policy Title"][i]))

    #         driver.find_element_by_name("sNamedInsMSC").clear()
    #         driver.find_element_by_name('sNamedInsMSC').send_keys("Certified Mail")
            
    #         '''                         Named Insured Line 1                '''

    #         driver.find_element_by_name("sNamedInsName1").clear()
    #         driver.find_element_by_name("sNamedInsName1").send_keys(str(df["Named Insured Line 1"][i]))

    #         ''''--------------ADDRESS LINE 1-------------'''
    #         driver.find_element_by_name("sAddr1").clear()
    #         driver.find_element_by_name("sAddr1").send_keys(str(df["Address Line 1"][i]))

    #         '''----------------CITY--------------------'''
    #         driver.find_element_by_name("sCity").clear()
    #         driver.find_element_by_name("sCity").send_keys(str(df["City"][i]))

    #         driver.find_element_by_name("sState").send_keys(str(df["State1"][i]).strip())
            
    #         if len(str(df["Zip Code"][i]).strip())==5:
    #             driver.find_element_by_name("sZip").clear()
    #             driver.find_element_by_name("sZip").send_keys(str(df["Zip Code"][i]))
    #         else:
    #             driver.find_element_by_name("sZip").clear()
    #             driver.find_element_by_name("sZip").send_keys("0"+str(df["Zip Code"][i]))
        
    #         '''---------------NAIC CODE-------------'''
        
    #         if str(df["NAIC Number"][i])=="19518":
    #             driver.find_element_by_name("iNAICID").send_keys("Greenwich Insurance Company | 19518 | 0")
    #         elif str(df["NAIC Number"][i])=="37885":
    #             driver.find_element_by_name("iNAICID").send_keys("XL Specialty insurance Co. | 37885 | 0")
    #         elif str(df["NAIC Number"][i])=="24554":
    #             driver.find_element_by_name("iNAICID").send_keys("XL Insurance America, Inc. | 24554 | 0")
    #         elif str(df["NAIC Number"][i])=="36940":
    #             driver.find_element_by_name("iNAICID").send_keys("Indian Harbor Insurance Co. | 36940 | 0")    
            
    #         '''-------------BRANCH NAME------------'''
    #         nip=Select(driver.find_element_by_name("BranchName"))
    #         nip.select_by_index(3)
    #         producernumber=driver.find_element_by_name("BranchName")


    #         '''--------------PRODUCER NUMBER---------'''

    #         for missing_data in range(1,5):
    #             time.sleep(1)
    #             pcode=driver.find_element_by_name("sProducerNbr").get_attribute('value')
    #             print(pcode)
    #             if pcode=="001":
    #                 break
    #             else:    
    #                 driver.find_element_by_name("sProducerNbr").clear()
    #                 driver.find_element_by_name("sProducerNbr").send_keys('001',Keys.TAB)
                    
                
                
    #         '''---------------INCEPTION DATE-----------'''
    #         driver.find_element_by_name("dtInception").clear()
            
    #         inception=str(df["Inception Date"][i]).split(" ")[0].split("-")
    #         inception_date=str(inception[1])+"/"+str(inception[2])+"/"+str(inception[0])
            
    #         driver.find_element_by_name("dtInception").send_keys(inception_date)

    #         '''---------------EXPIRATION DATE-----------'''

    #         driver.find_element_by_name("dtExpiration").clear()
    #         exception=str(df["Expiration Date"][i]).split(" ")[0].split("-")
    #         exception_date=str(exception[1])+"/"+str(exception[2])+"/"+str(exception[0])
    #         driver.find_element_by_name("dtExpiration").send_keys(exception_date)


    #         '''------TOTAL $-------------------------'''

    #         if str(df["State"][i])=="NJ":
    #             try:
    #                 driver.find_element_by_name("curPremiumDue").clear()
    #             except:
    #                 print("Clear")
                    
    #             try:
    #                 driver.find_element_by_name("curPremiumDue").send_keys("{:,}".format(int(str(df["Total $"][i]))))
    #             except:
    #                 driver.find_element_by_name("curPremiumDue").send_keys(str(df["Total $"][i]).split('.')[0])

    #                 '''--------------DUE DATE-------------------'''

    #             desc_test1=str(df["AgeDate"][i]).split(" ")[0].split("-")
    #             desc_test1=str(desc_test1[1])+"/"+str(desc_test1[2])+"/"+str(desc_test1[0])

    #             try:
    #                 driver.find_element_by_name("dtPremium").clear()
    #             except:
    #                 print("Clear")
    #             driver.find_element_by_name("dtPremium").send_keys(desc_test1)

    #         driver.switch_to.default_content()
            
    #         driver.switch_to.frame("LeftFrame")

    #         '''-------------------CREATE NOTICE----------'''
    #         driver.find_element_by_link_text('Create Notice').click()
                        
    #         '''-------------TEXAS CONDITION----------'''
    #         while True:
    #             try:
    #                 driver.find_element_by_partial_link_text("Continue").click()
    #             except:
    #                 print("Texas")
    #             try:
    #                 driver.find_element_by_partial_link_text("Cancellation").click()
    #                 break
    #             except:
    #                 pass
    #                 # print("Click Cancellation again")
            
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("RightFrame")

    #         '''-----------------REASON CODE------------'''
    #         time.sleep(3)
    #         try:
    #             select1 = driver.find_element_by_name(
    #                     "sReasonID").send_keys("NONPAYMENT OF PREMIUM")
    #         except:
    #             print("Select")
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("LeftFrame")
    #         driver.find_element_by_partial_link_text("Continue").click()
    #         time.sleep(2)


    #         '''----------------PREMIUM------------------'''

    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("RightFrame")
    #         k=j=1
    #         try:
    #             try:
    #                 driver.find_element_by_name('curPremiumDue').send_keys("{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
    #             except:
    #                 driver.find_element_by_name('curPremiumDue').send_keys((str(df["Total $"][i]).split(".")[0]))
    #             k=0
    #         except:
    #             print("Current Premium")


    #         try:
    #             duedate=str(df["Due Date"][0]).split(" ")[0].split("-")
    #             duedate_final=str(duedate[1])+"/"+str(duedate[2])+"/"+str(duedate[0])
    #             driver.find_element_by_name('dtPremium').send_keys(duedate_final)
    #             j=0
    #         except:
    #             print("Current Premium")



    #         '''---------------NEWYORK CONDITION----------------------------'''

    #         if k==0 or j==0:
    #             driver.switch_to.default_content()
    #             driver.switch_to.frame("LeftFrame")
    #             driver.find_element_by_partial_link_text("Continue").click()
                    
    #             driver.switch_to.default_content()
    #             driver.switch_to.frame("RightFrame")

    #         try:
    #             if df["State"] == 'NY':
    #                 driver.find_element_by_name('p295073_S').send_keys("{:}".format(int(str(df["Total $"][i]).split(".")[0])))
    #             else:
    #                 driver.find_element_by_name('p295073_S').send_keys("{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
    #         except:
    #             try:
    #                 driver.find_element_by_name('p295073_S').send_keys(str(df["Total $"][i]).split(".")[0])
    #             except:
    #                 print("Current Premium")
                
    #         try:
    #             driver.find_element_by_name('p272').send_keys("800-446-7647 \n NIP Group, Inc. \n 900 Route 9 North \nWoodbridge NJ 07095")
    #         except:
    #             print("Address Phone")

    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("LeftFrame")
            

    #         '''----------------CLICKING CONTINUE------------------'''

    #         while True:
    #             time.sleep(1)
    #             question=driver.find_element_by_xpath('//*[@id="menuDots"]/div[1]/form/table/tbody/tr[1]/td').text
    #             if question=="Reason Details":
    #                 break
    #             else:
    #                 driver.switch_to.default_content()
    #                 driver.switch_to.frame("LeftFrame")
    #                 driver.find_element_by_partial_link_text("Continue").click()
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("RightFrame")
            
    #         time.sleep(5)
            


    #         '''----------------------DESCRIPTION-------------------------'''
    #         from datetime import datetime as dt
    #         checkDate = dt.strptime('05-15-2022','%m-%d-%Y')
    #         reasonDetails = '''due to non-payment of premium of $500 which was due on 02/01/2020

    # '''
    #         for desc in range(len(df1)):
    #             if str(df["State"][i]) in str(df1["State"][desc]) or str(df1["State"][desc]) in str(df["State"][i]):
    #                 if df['Inception Date'][i] < checkDate:
    #                     date_desc=str(date.today()+ timedelta(days=25)).split("-")
    #                     date_desc=str(date_desc[1])+"/"+str(date_desc[2])+"/"+str(date_desc[0])
    #                     descr=str(df1["Terms and Description"][desc]).replace("05/07/2020", date_desc)
    #                     desc_test=str(df["AgeDate"][i]).split(" ")[0].split("-")
    #                     desc_test=str(desc_test[1])+"/"+str(desc_test[2])+"/"+str(desc_test[0])
    #                     expiration=descr.replace("02/01/2020", desc_test)
    #                     try:
    #                         premium=expiration.replace("500","{:,}".format(int(str(df["Total $"][i]).split(".")[0])))
    #                     except:
    #                         premium=expiration.replace("500",(str(df["Total $"][i]).split(".")[0]))
    #                     driver.find_element_by_name("mReasonDesc").send_keys(premium)
    #                     break
    #                 elif df['Inception Date'][i] >= checkDate:
    #                     driver.find_element_by_name("mReasonDesc").send_keys(reasonDetails)

    #         from datetime import datetime
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("LeftFrame")
    #         driver.find_element_by_partial_link_text("Continue").click()
    #         time.sleep(5)
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("RightFrame")


    #         '''----------------------MAIL DATE-----------'''
            
    #         try:
    #             driver.find_element_by_name('dtMail_D').clear()
    #             date_t=str(date.today()+ timedelta(days=1)).split("-")
    #             date_final=str(date_t[1])+"/"+str(date_t[2])+"/"+str(date_t[0])
    #             driver.find_element_by_name('dtMail_D').send_keys(date_final)
    #         except:
    #             date_t=str(date.today()+ timedelta(days=1)).split("-")
    #             date_final=str(date_t[1])+"/"+str(date_t[2])+"/"+str(date_t[0])
    #             try:
    #                 driver.find_element_by_name('dtMail_D').send_keys(date_final)
    #             except:
    #                 print("No Mailing Date")
                
            
            
                
    #         print(str(int(df["Days"][i])-5))
    #         try:
    #             driver.find_element_by_name('iMailLeadDays_N').clear()
    #             driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])))
    #         except:
    #             try:
    #                 driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])))
    #             except:
    #                 try:
    #                     driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])))
    #                 except:
    #                     driver.find_element_by_name('iMailLeadDays_N').send_keys(str(int(df["Days"][i])))
                        
    #         try:
    #             driver.switch_to_alert().accept()
    #         except:
    #             print("Alert")   
             
            
    #         for desc in range(len(df1)):
    #             if str(df["State"][i]) in str(df1["State"][desc]):
    #                 try:
    #                     driver.find_element_by_name('iAdvDays_N').clear()
    #                     driver.find_element_by_name('iAdvDays_N').send_keys(str(int(df1["Advance Days"][desc])))
    #                 except:
    #                     try:
    #                         driver.find_element_by_name('iAdvDays_N').send_keys(str(int(df1["Advance Days"][desc])))
    #                     except:
    #                         try:
    #                             driver.find_element_by_name('iAdvDays_N').send_keys(str(int(df1["Advance Days"][desc])))
    #                         except:
    #                             try:
    #                                 driver.find_element_by_name('iAdvDays_N').send_keys(str(int(df1["Advance Days"][desc])))
    #                             except:
    #                                 driver.find_element_by_name('iAdvDays_N').send_keys(str(int(df1["Advance Days"][desc])))
            
    #         try:
    #             driver.switch_to_alert().accept()
    #         except:
    #             print("certifi")
    #             print("Error")
                
    #         try:
    #             driver.find_element_by_name('iMailTypeID').clear()
    #             driver.find_element_by_name('iMailTypeID').send_keys("Certified")
    #         except:
    #             try:
    #                 driver.find_element_by_name('iMailTypeID').send_keys("Certified")
    #             except:
    #                 driver.find_element_by_name('iMailTypeID').send_keys("Certified")
        
    #         try:
    #             driver.switch_to_alert().accept()
    #         except:
    #             print("Pop Up")
    #         try:
    #             driver.find_element_by_name('sAsOfPhrase').send_keys("12:01 A.M. Local Time")
    #         except:
    #             driver.find_element_by_name('sAsOfPhrase').send_keys("12:01 A.M. Local  Time")
    #         driver.switch_to.default_content()
    #         driver.switch_to.frame("LeftFrame")
    #         driver.find_element_by_partial_link_text("Continue").click()
            
    #         time.sleep(5)
    #         driver.find_element_by_partial_link_text("Generate").click()
    #         WebDriverWait(driver, 50).until(EC.number_of_windows_to_be(2))
    #         time.sleep(3)
            

    #         driver.quit()
    #         print(df["State"][i])

    #         '''------------RENAME FILE-----------'''    
    #         for file in os.listdir(r"C:\NIP NOC\05-10 CR Testing\PDF Rename test"):
    #             print('Renaming the pdf...')
    #             if file.endswith(".pdf"):
    #                shutil.move(r"PDF Rename test/" + str(file), r"PDF Download/" + str(df["Policy Number"][i]) + str(df["Account"][i]) + str(df["State"][i]) + ".pdf")
    #         df["Status"][i]="Generated" 
    #     except:
    #         Pnum=df["Policy Number"][i]
            
    #         df["Status"][i]="Not Generated"    
            
            
        
            #error_mail(Pnum)
    df.to_excel(r"C:\NIP NOC\output.xlsx",index=False)
    

merge_pdf_by_account_id()
#completed_mail()



