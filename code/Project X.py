# -*- coding: utf-8 -*-
"""
@author: Vinay Krishna
"""

#Important libraries required to work with selenium and pandas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from zipfile import ZipFile
import PyPDF2
import pandas as pd

#Intialsing the webdriver for chrome
chromeOptions=webdriver.ChromeOptions()
#Allowing the downloads from chrome to be save on default directory.
prefs={"download.default_directory":"C:\PROJECT X"}
chromeOptions.add_experimental_option("prefs",prefs)
#Path to your chrome driver executable file
driver= webdriver.Chrome(r'C:\PROJECT X\chromedriver_win32\chromedriver.exe',chrome_options=chromeOptions)
#Logging into Gmail using get method and performing the actions using Action Chains
driver.get("https://www.gmail.com/")
action=ActionChains(driver)
"""
1. When gmail login page is opened you can inspect the page by going to Inspect tab on chrome.
2. Find element by name finds the name given for that particular boxes displayed like username/password space
3. After entering the credentials, find element by class name is used to hover your mouse on the next/enter button to click and proceed
4. Once login is completed, a mail with particular subject is found by the Xpath. The Xpath can be extracted from the inspect tab on Chrome and finding the relative xpath for your email with a specific or unique body name.
"""
elem=driver.find_element_by_name("identifier")
elem.send_keys("") #enter your email address
driver.find_element_by_class_name("CwaK9").click()
time.sleep(2)
elem=driver.find_element_by_name("password")
elem.send_keys("") #enter your password
driver.find_element_by_class_name("CwaK9").click()
time.sleep(20)
elem=WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//tr//td//div//span[@class='bog']/span[contains(text(),'Customer Details')]")))
time.sleep(3)
elem.click() # clicking the email to open
time.sleep(2)
v = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='asa']/div[contains(@class,'aZh')]"))) #downloading the files or zip folder using xpath and directly allowing to click on the icon for download
time.sleep(3)
action.move_to_element(v).click().perform()
time.sleep(15)
driver.quit()

"""
1. Using Zipfile library unzipping multiple pdfs downloaded and saving into project X folder.
2. Using a for loop to iterate through all pdf's present in the folder and reading using Pdf reader library.
3. Using string manipulation techniques to extract data from different columns in a pdf. You can call the function below as many times as you required based on the number of columns in your pdf
4. Using Excel functions from Pandas data frame to create a single excel file containing the customer details from all pdfs in structured format.
"""

with ZipFile(r"C:\PROJECT X\customerdetails.zip",'r') as zipObj:
    zipObj.extractall(r"C:\PROJECT X")


text = " " 
b1 = []
b2 = []
for i in range(0,10):
    q=open(r"C:\PROJECT X\Customer Details {0}.pdf".format(i),'rb')
    pdfReader = PyPDF2.PdfFileReader(q)
    pageObj = pdfReader.getPage(0) 
    texti= pageObj.extractText()
    yi=texti.replace("\n","") #Replacing \n values with empty string
    #creating a function to get the string for the desired title
    def get_text_btw(string,start,end):
         start_index=string.index(start)+len(start)
         z=string[start_index:string.index(end)]
         if(z=='' or z==' ' or z=='  ' or z=='   '):
             z='NA'
             return (z)
         else:
             r=z.strip()
             return (r)
    a1 = get_text_btw(yi,"Name: ","Age: ") #calling the function to extract the Name of person which lies between Name and age text in pdf
    a2 = get_text_btw(yi,"Age: ","Interests: ")
    b1.append(a1)
    b2.append(a2)
#creating a dataframe using pandas library and assiging data frame values with respective keys and values as list            
dataframe = pd.DataFrame({'Name':b1,'Age':b2,'Interests':b3}) 
#creating an excel sheet
writer_object = pd.ExcelWriter("C:\PROJECT X\X.xlsx",engine ='xlsxwriter')
#interchanging the rows and columns in the excel sheet
z=dataframe.transpose()
z.to_excel(writer_object, sheet_name ='1')
worksheet_object  = writer_object.sheets['1']
worksheet_object.set_column('A:F', 30)  #No of columns you wanted on th excel sheet.
writer_object.save() 
