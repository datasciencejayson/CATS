# -*- coding: utf-8 -*-
"""
Created on Tue Jan 30 09:28:45 2018

@author: backesj
"""

import sys
import os

os.system('mode con: cols=100 lines=50')
  
print("""   
         
********************************************************************
*                      Welcome to CATS                             *
*            A PI tool to find Bene phone numbers                  *
********************************************************************
    Created by: Jayson Backes
    Contact:    backesj@admedcorp.com


                                _.----.
                      |\---/|  / ) cat|
          ------------;     |-/ /|food|---
                      )     (' / `----'
          ===========(       ,'==========
          ||   _     |      |
          || o/ )    |      | o
          || ( (    /       ;              (George)
          ||  \ `._/       /
          ||   `._        /|
          ||      |\    _/||
        __||_____.' )  |__||____________
         ________\  |  |_________________
                  \ \  `-.
                   `-`---'  
             
                            _,'|             _.-''``-...___..--';)
                           /_ \'.      __..-' ,      ,--...--'''
        (Kimchi)          <\    .`--'''       `     /'
                           `-';'               ;   ; ;
                     __...--''     ___...--_..'  .;.'
                    (,__....----'''       (,..--''

                   and (Kirra)
            
                   .-o=o-.
               ,  /=o=o=o=\ .--.
              _|\|=o=O=o=O=|    |
          __.'  a`\=o=o=o=(`\   /
          '.   a 4/`|.-""'`\ \ ;'`)   .---.
            \   .'  /   .--'  |_.'   / .-._)
             `)  _.'   /     /`-.__.' /
              `'-.____;     /'-.___.-'
                       `"'`
             """)
          
input(' Would you like to run the CATS report, Y or N: ')

print("")
print(" **** Running Program ****")
print(" Please wait for the prompt")

#import packages

from tkinter import *
import pandas as pd
from tkinter import messagebox 
from pathlib import Path
import re
from sqlalchemy import create_engine # database connection
import difflib
import requests
from bs4 import BeautifulSoup as bs4
import time
from selenium import webdriver
#from selenium.webdriver.support.ui import Select
#from selenium.webdriver.common.keys import Keys

# start tkinter module for hic and location entry
def tkinterNums():
    """
    this function will ask the user to enter the value of HIC or SSN as well
    as the location of the xlsx file they want to import. Only xlsx will work.
    the function then imports the list and exits the tkinter module. There are
    no parameters to use the function. All parameters are enter by the user.
    """
    #define global variables
    global master, nums
    
    # this function that exits the tkinter box if the user chooses
    def quitTkinterBox():
        global master, nums, numType, file
        # destroys the tkinterbox
        master.destroy()  
        # returns nums and numType as strings of none
        nums = 'None'
        numType = 'None'
        return nums, numType, file
    
       
    def importNums():
        """ this function is the main function that imports the data.
        it is called from the validateEntries function after the user clicks
        the run button on the tkinter box. there are no parameters. 
        all parameters are set by the user
        """
        
        # set the global parameters
        global master, nums, numType, file
        # get the numType (HIC or SSN) the the user entered 
        numType = e1.get().upper()
        # get the file path that the user entered
        file = e2.get()
#        print(file)
        # create pandas table "nums" from the excel file
        nums = pd.read_excel("%s" % file, header=None)
        # destroy tkinter box
        master.destroy()   
        # return results
        return nums, numType, file
        
    def validateEntries():
        # define valid list types
        validEntries = ['hic', 'ssn']
        global master, nums
        # get the numType (HIC or SSN) the the user entered 
        text = e1.get()
        # set myFile pointer to file path
        myFile = Path("%s" % e2.get())
        # validate that the file exists and is numType is correct
        if text.lower() in validEntries and myFile.is_file():
            importNums()
            return True
        # go through modalites to construct the type of message to present 
        # to the user.
        else:
            if text.lower() not in validEntries and myFile.is_file():
                
                messagebox.showinfo("*** Warning ***", "Please decalre entry type: HIC or SSN")
#            e1.focus()
#            e1.delete(0, END)
            
                return False
            elif text.lower() in validEntries and myFile.is_file() == False:
                messagebox.showinfo("*** Warning ***", "The file you entered does not exist, \n please try again.")
                return False
            elif text.lower() not in validEntries and myFile.is_file() == False:
                messagebox.showinfo("*** Warning ***", "Please decalre entry type: HIC or SSN \n Also, the file does not exist.")
                return False
                
#    def validateFile(text):
#        global master, hics
#
#        my_file = Path("%s" % text)
#        if my_file.is_file(): return True
#        else:
#            messagebox.showinfo("*** Warning ***", "The file you entered does not exist, \n please try again.")
#            e2.focus()
#            return False

            
    # this is the tkinter code that presents the tkinter box to the user
    
    master = Tk()
    master.title("Enter Type of Input and location of file")
    master.lift()
    
    master.attributes('-topmost',True)
    
    master.after_idle(master.attributes,'-topmost',False)
    Label(master, text="Enter the Type of Bene Number in your Excel File: ").grid(row=0, column=0)
    Label(master, text="Example: HIC or SSN").grid(row=1, column=0, sticky=W)
    
    Label(master, text="-------------------------------------------------------------------------------------------").grid(row=2, columnspan=2, sticky=W)


    Label(master, text="Enter the Full File Path & name of your Excel File: ").grid(row=3, column=0)
    Label(master, text="Example: H:/Work/hic_list.xlsx").grid(row=4,column=0, sticky=W)
    
    e1 = Entry(master)
    e1.grid(row=0, column=1)
#    vcmd2 = (master.register(validateFile), "%P")
#    e2 = Entry(master, validate = "focusout", validatecommand=vcmd2)
    e2 = Entry(master)
    e2.grid(row=3, column=1)
    Label(master, text="-------------------------------------------------------------------------------------------").grid(row=5, columnspan=2, sticky=W)

    Button(master, text='Run', command=validateEntries).grid(row=6, column=0)
    Button(master, text='Quit', command= quitTkinterBox).grid(row=6, column=1)

    master.mainloop()    
    return nums, numType, file


nums, numType, file = tkinterNums()


# check that an actual df was returned. If string, somthing when wrong, so
# exist program

def dfChecker():

    def exitProgram():
        print('raise SystemExit')
    def warning():   
        exitVar = input("Something went wrong with the program. Enter X to exit, and \
              verify that your excel file is an xlsx file and that it is \
              comprised of HICs or SSNs: ")
        if exitVar in ['x', 'X']:
            exitProgram()
        else:  
            print("You didn't enter 'x' or 'X', please enter x or X to exit")
            time.sleep(1)
            warning()

    def retry():
        exitVar = input("Please Enter X to exit, and verify that your \
                        excel file is an xlsx file and that it is \
                        comprised of HICs or SSNs: ")
        if exitVar in ['x', 'X']:
            exitProgram()
        else:  
            print("You didn't enter 'x' or 'X', please enter x or X to exit")
            time.sleep(1)
            warning()

    def dfCheck():
        if str(type(nums)) != "<class 'pandas.core.frame.DataFrame'>":
            warning()
        else:
            print("\n STEP 1: Pulling Bene Information")
    dfCheck()
 
    
dfChecker()
    
    

tempList = list(nums[0])


    
beneList = []
for i in tempList:
    if bool(re.search(r'\d', i)) == True:
        beneList.append(i)
 
beneTuple = tuple(beneList)      



#df = pd.read_sql_query("SELECT * FROM beneTable", smallDb)


# create an engine that will mangae our bene data

beneDatabase = create_engine('sqlite:///F:\\backesj\\fullBeneDatabase.db')

#beneDatabase = create_engine('sqlite:///F:\\backesj\\smallBeneDatabase.db')
#deadDatabase = create_engine('sqlite:///F:\\backesj\\tempBeneDatabase.db')

# create an engine that will mangae our facility data

facDatabase = create_engine('sqlite:///F:\\backesj\\facAddressDatabase.db')

deadDatabase = create_engine('sqlite:///F:\\backesj\\deadBeneDatabase.db')
    
    

if numType == 'HIC': 

    print('\nUsing HIC list to pull HIC information from SAS databases, \n this should take approximatly 10-20 minutes.')
    
    dfDeadBene = pd.read_sql_query("SELECT * FROM deadTable where HIC in %s" % str(beneTuple), deadDatabase)
    
    dfBene = pd.read_sql_query("SELECT * FROM beneTable where HIC in %s" %  str(beneTuple), beneDatabase)
    
    deadBenes = list(dfDeadBene['HIC'])


else:

    print('\n Using SSN list to pull SSN information from SAS databases, \n this should take approximatly 10-20 minutes.')
    dfDeadBene = pd.read_sql_query("SELECT * FROM beneTable where recip_ssn in %s" % str(beneTuple), deadDatabase, index_col='index')
     
    dfBene = pd.read_sql_query("SELECT * FROM beneTable where recip_ssn in %s and recip_ssn not in %s " %  (str(beneTuple), str(deadBenes)), beneDatabase, index_col='index')
    
    deadBenes = list(dfDeadBene['recip_ssn'])
    
    
    
# add search name and index
beneFullNameList = []
for index, row in dfBene.iterrows():
    if str(row['RECIP_NAME_MIDDLE_INIT']) == 'None':
        beneFullNameList.append(str(row['RECIP_NAME_FIRST']) + ' ' + str(row['RECIP_NAME_LAST']))
    else:
        beneFullNameList.append(str(row['RECIP_NAME_FIRST']) + ' '+ str(row['RECIP_NAME_MIDDLE_INIT']) + ' ' + str(row['RECIP_NAME_LAST']))

dfBeneFullName = pd.DataFrame(beneFullNameList)
dfBeneFullName.columns = ['FULL_NAME']

dfBene2 = pd.concat([dfBene,dfBeneFullName], axis = 1) 

dfBene2['BENE_IDX'] = range(1, len(dfBene2) +1)

dfBene2['SEARCH_STATE'] = ' ' + dfBene2['recip_state'].map(str)
# get possible facility information based on zip code match

# create a list of ZIP codes to search fac table
dfBeneZipList= list(dfBene2['ZIP5'])

dfBeneZipList2 = [x for x in dfBeneZipList if type(x) != float]

dfBeneZipTuple = tuple(dfBeneZipList2)

dfFac = pd.read_sql_query("SELECT * FROM facTable where ZIP5 in %s" % str(dfBeneZipTuple), facDatabase)

# add index to fac table
dfFac['FAC_IDX'] = range(1, len(dfFac) +1)


# create bene table lists
dfBeneIndexList = list(dfBene2['BENE_IDX'])
dfBeneAddList = list(dfBene2['ADDRESS'])
dfBeneNameList = list(dfBene2['FULL_NAME'])
dfBeneZipList = list(dfBene2['ZIP5'])
dfBeneStateList = list(dfBene2['recip_state'])
dfBeneSearchStateList = list(dfBene2['SEARCH_STATE'])
dfBeneHicList = list(dfBene2['HIC'])
dfBeneSsnList = list(dfBene2['recip_ssn'])

#dfBenePhoneList = [x for x in dfBene['TELEPHONE'].map(str) if x != 'None']

# create new varible PHONE that includes all phone number information
dfFac['PHONE'] = list(dfFac['TELEPHONE'].map(str).replace('None','')  + dfFac['TELEPHONE1'].map(str).replace('None','') + \
    dfFac['TELEPHONE2'].map(str).replace('None',''))

# create fac table lists
dfFacPhoneList = list(dfFac['PHONE'])
dfFacIndexList = list(dfFac['FAC_IDX'])
dfFacAddList = list(dfFac['ADDRESS'])
dfFacZipList = list(dfFac['ZIP5'])

# fuzzy search fac table for possible phone numbers
resultDict = {}
for i, ivalue in enumerate(dfBeneAddList):
    for j, jvalue in enumerate(dfFacPhoneList):
        if dfBeneZipList[i] == dfFacZipList[j] and dfBeneZipList[i] != None:
            result = difflib.SequenceMatcher(None, dfBeneAddList[i], dfFacAddList[j]).ratio()
            if result > .8:
                resultDict[i] = [dfBeneIndexList[i], \
                   dfBeneAddList[i], dfFacAddList[j], dfBeneNameList[i], dfFacPhoneList[j].strip()]

dfFacFinal = pd.DataFrame()
facNameList = []
facPhoneList = []
for key, value in resultDict.items():
    facPhoneList.append(value[4])
    facNameList.append(value[3])

facPhoneDf = pd.DataFrame({'MATCH_NAME':facNameList, 'Facility Phone':facPhoneList})    
    
# create bene Dictionary to house all pertinent information
def createBeneDict():
    beneDict = {}
    for i, ivalue in enumerate(dfBeneIndexList):
        # check to see if any name, state combination is currently in the dict
        beneDict[i] = [dfBeneIndexList[i], dfBeneNameList[i].replace('None ',''), \
                 dfBeneStateList[i], dfBeneHicList[i], dfBeneSsnList[i], \
                 dfBeneSearchStateList[i]]
    return beneDict

beneDict = createBeneDict()


print("\n STEP 2: Webscraping")

# create url variable of web address
url='https://www.truepeoplesearch.com/'

# prepare df for webscrape

print('\n Starting truepeoplesearch.com')

#zone = [', AL',', AR',', GA',', LA',', MS',', NC',', TN',', WV']

# could possibly add an area that would emcumpass possible resutls
#area = ['AL','AR','GA','LA','MS','NC','TN','WV']

# define dedup function to dedup lists
def dedup(seq):
    """
    removes duplicate values from a list or 
    duplicate characters from a string 
    """
    if type(seq) == list:
        seen = set()
        seen_add = seen.add
        return [x for x in seq if not (x in seen or seen_add(x))]    
    elif type(seq) == str:
        seen = set()
        seen_add = seen.add
        return ''.join([x for x in seq if not (x in seen or seen_add(x))])    
    else:
        print("Currently function can only handle lists and strings")

# create empty dictionary and list  
# ps, hardest webscraped i have ever done. the captch problem makes it even worse
fullDict = {}
nameList = []
for key, value in beneDict.items(): 
    nameDict = {}
    counter = -1
    if value[1] not in nameList:
        print('Searching TruePeopleSearch for %s'%  value[1])
        nameList.append(value[1])
        splitName = '%20'.join(value[1].split())
    #    for k, kvalue in enumerate(replaceList)
        cleanName = value[1]
        first = value
    #    splitName = 'john%20smith'
#        print(splitName)
        flag = 'Y'
        pageCount = 0
        while flag == 'Y':
#            print(flag)
            pageCount +=1
            tempURL = url+'results?name='+splitName+'&citystatezip=%s&page=%s' % (value[2], pageCount)
            pageContent = requests.get(tempURL).content
    #        time.sleep(1)
#            print(tempURL)
            if 'captchasubmit?returnUrl' in str(pageContent): 
#                print('first')    
                while 'captchasubmit?returnUrl' in str(pageContent):
                
                    print('captcha found')
                    def afterCaptcha():
                        global captcha, pageContent
                        pageContent = requests.get(tempURL).content
                        print("Program will continue until another captcha is requested")
                        captcha.destroy()
                        
                    def quitTkinterBox():
                        global captcha
                        captcha.destroy()
                        print("raise SystemExit")
                    
                    captcha = Tk()
                    captcha.lift()
                    
                    captcha.attributes('-topmost',True)
                    
                    captcha.after_idle(captcha.attributes,'-topmost',False)
    
#                    messagebox.showinfo("*** Warning ***", "The website has requested a captcha \n please go to https://www.truepeoplesearch.com/ and manually \n  solve the captcha")
                    
                    Label(captcha, text="*** Warning ***, The website has requested a captcha \n please go to https://www.truepeoplesearch.com/ and manually \n solve the captcha. When you have finished, \n come back to this window and please press continue.").grid(row=0, sticky = W)
                    
                    Button(captcha, text='Continue', command=afterCaptcha).grid(row=4, column=1, sticky=W, pady=1)
                    
                    Button(captcha, text='Quit Program and Exit', command= quitTkinterBox).grid(row=4, column=2, sticky=W, pady=4)
                
                    captcha.mainloop()  
                    
                print("Continuing")
                pageContent = requests.get(tempURL).content
    #                time.sleep(1)
                soup = bs4(pageContent, "html.parser")
#                print(soup)
                linkList = []
                diffList = []
                if str(soup).find('btnNextPage') == -1:
                    flag = 'N'
                for card in soup.find_all(attrs= {'class':'card card-block shadow-form card-summary'}):
                    if str(value[5]) in card.text: 
                        for h4 in card.find_all(attrs= {'class':'h4'}):
#                            print(value[1])
#                            print(h4.text.strip().upper())
                            result = difflib.SequenceMatcher(None, value[1], h4.text.strip().upper()).ratio()
#                            print(result)
                        for a in card.find_all('a'):
                            if 'name' in a['href']:
                                if 'page' not in a['href']:
                                    if a['href'] not in linkList:
    #                                    if result > .5:
                                            diffList.append(result)
                                            linkList.append(a['href'])
                    else:
                        None
                    
                        
            else:
                soup = bs4(pageContent, "html.parser")
#                print(soup)
                linkList = []
                diffList = []
                if str(soup).find('btnNextPage') == -1:
                    flag = 'N'
                for card in soup.find_all(attrs= {'class':'card card-block shadow-form card-summary'}):
                    if str(value[5]) in card.text: 
                        for h4 in card.find_all(attrs= {'class':'h4'}):
#                            print(value[1])
#                            print(h4.text.strip().upper())
                            result = difflib.SequenceMatcher(None, value[1], h4.text.strip().upper()).ratio()
#                            print(result)
                        for a in card.find_all('a'):
                            if 'name' in a['href']:
                                if 'page' not in a['href']:
                                    if a['href'] not in linkList:
    #                                    if result > .5:
                                            diffList.append(result)
                                            linkList.append(a['href'])
                    else:
                        None
    
            for i, ivalue in enumerate(linkList):
                counter += 1
                infoDict = {}
                tempURL = url+linkList[i]
                pageContent2 = requests.get(tempURL).content
    #            time.sleep(1)
                if 'captchasubmit?returnUrl' in str(pageContent2): 
#                    print('first')
                    while 'captchasubmit?returnUrl' in str(pageContent2):
                    
                        print('captcha found')
                        def afterCaptcha():
                            global captcha, pageContent2
                            pageContent2 = requests.get(tempURL).content
    #                        time.sleep(1)
            #                print("Failed to solve captcha. Ending program. Please try again.")
                            print("Program will continue until another captcha is requested")
                            captcha.destroy()
                            
                        def quitTkinterBox():
                            global captcha
                            captcha.destroy()
                            print("raise SystemExit")
                        
                        captcha = Tk()
                        captcha.lift()
                        
                        captcha.attributes('-topmost',True)
                        
                        captcha.after_idle(captcha.attributes,'-topmost',False)
        
#                        messagebox.showinfo("*** Warning ***", "The website has requested a captcha \n please go to https://www.truepeoplesearch.com/ and manually \n  solve the captcha")
                    
                        Label(captcha, text="*** Warning ***, The website has requested a captcha \n please go to https://www.truepeoplesearch.com/ and manually \n solve the captcha. When you have finished, \n come back to this window and please press continue.").grid(row=0, sticky = W)
                    
                        
                        Button(captcha, text='Continue', command=afterCaptcha).grid(row=4, column=1, sticky=W, pady=1)
                        
                    #    Button(master, text='Run', command=importNums).grid(row=4, column=1, sticky=W, pady=1)
                        Button(captcha, text='Quit Program and Exit', command= quitTkinterBox).grid(row=4, column=2, sticky=W, pady=4)
                    
                        captcha.mainloop()  
                        
                    print("Continuing")
                    soup = bs4(pageContent2, "html.parser")
                    phoneList = []
#                    print(ivalue)
                    infoDict['name'] = soup.find(attrs= {'class','h2'}).text.strip()
                    infoDict['age'] = soup.find(attrs= {'class','content-value'}).text.strip()
                    infoDict['address'] = soup.find(attrs= {'link-to-more','link-to-more'}).text.strip()
                    infoDict['match'] = value[0]
                    infoDict['origFullName'] = value[1]
                    for a in soup.find_all('a'):
                        if 'phoneno' in a['href']:
                            phone = a['href'][a['href'].find('=')+1:]
                            if phone not in phoneList:
                                phoneList.append(phone)
                    infoDict['phone'] = phoneList
                    infoDict['source'] = 'TPS'
                    infoDict['diff'] = diffList[i]

    
                else:
                    
                    soup = bs4(pageContent2, "html.parser")
                    phoneList = []
#                    print(ivalue)
                    infoDict['name'] = soup.find(attrs= {'class','h2'}).text.strip()
                    infoDict['age'] = soup.find(attrs= {'class','content-value'}).text.strip()
                    infoDict['address'] = soup.find(attrs= {'link-to-more','link-to-more'}).text.strip()
                    infoDict['match'] = value[0]
                    infoDict['origFullName'] = value[1]

                    for a in soup.find_all('a'):
                        if 'phoneno' in a['href']:
                            phone = a['href'][a['href'].find('=')+1:]
                            if phone not in phoneList:
                                phoneList.append(phone)
                    infoDict['phone'] = phoneList
                    infoDict['source'] = 'TPS'
                    infoDict['diff'] = diffList[i]
                nameDict[counter] = infoDict
            fullDict[value[0]] = nameDict 
    else:
        None
        
################# THAT'S THEM ################################



"""
Created on Thu Dec 21 13:43:57 2017

That's Them written by saulnk, edited by backesj

"""


url='https://thatsthem.com/'

names=['Danny Drinkwater', 'Eden Hazard', 'Chelsea Soccer', 'James Smith']



thatThem=[]


thatsThemFinal = []
thatsThemDict = {}
nameList = []
print('\n Starting ThatsThem.com')
for key, value in beneDict.items():
    if value[1] not in nameList:
        time.sleep(1)
        nameList.append(value[1])
        name=str(value[1])
        print('Searching ThatsThem for %s' % name)
        # set driver
        driver = webdriver.PhantomJS(executable_path=r'S:\DA_work_files\DA_Work_Python\phantomjs-2.1.1-windows\bin\phantomjs.exe')
        time.sleep(1)

        driver.set_window_size(1124, 850) # set browser size.
        # use driver to get url
        time.sleep(1)

        driver.get(url)
        
        #Find the search box and input the name
        time.sleep(1)
        nameInput = driver.find_element_by_id('fullName')
        time.sleep(1)

        nameInput.send_keys(name)
        time.sleep(1)

        # click on search button
        submit = driver.find_element_by_css_selector('button.btn.btn-lg.btn-block').click()
        time.sleep(1)

         
        #get current page
        page_content = requests.get(driver.current_url).content
        time.sleep(1)

        soup = bs4(page_content, "html.parser")
        
        
        try:
            #finding the number or total results
            for line in soup.find_all("span",class_="ThatsThem-results-preheader"):
                results1 = ''.join(line.find_all(text=True))
                
            results=results1[7:10]
            results=int(results.strip())
            
            # Grab name, Address, Phone number, Email for each result
            name=[]
            address1=[]
            address2=[]
            address3=[]
            address4=[]
            phone=[]
            email=[]
            match=[]
            origFullName = []
            
            
            i=0
            for i in range(results):
                iteration= soup.find_all('div',class_="ThatsThem-record")[i]
                for span in iteration.find_all("span",itemprop='name'):
                    name1=span.text
                    name.append(name1)
                for span in iteration.find_all("span",itemprop='streetAddress'):
                    address1a=span.text
                    address1.append(address1a)
                for span in iteration.find_all("span",itemprop='addressLocality'):
                    address2a=span.text
                    address2.append(address2a)
                for span in iteration.find_all("span",itemprop='addressRegion'):
                    address3a=span.text
                    address3.append(address3a)
                for span in iteration.find_all("span",itemprop='postalCode'):
                    address4a=span.text
                    address4.append(address4a) 
                    
                #Phone Number
                num_phone=len(phone)
                for span in iteration.find_all("span",itemprop='telephone'):
                    phone1=span.text.replace('-','')
                    phone.append(phone1)
                
                if num_phone == len(phone):
                    no_num = 'No Number Found'
                    phone.append(no_num)
                else: 
                    pass
                    
                #Email Address
                num_email=len(email)
                for span in iteration.find_all("span",itemprop='email'):
                    email1=span.text
                    email.append(email1)
                
                if num_email == len(email):
                    no_mail = 'No Email Found'
                    email.append(no_mail)
                else: 
                    pass
                match.append(value[0])
                origFullName.append(value[1])
                
            # Change to data frames
            name=pd.DataFrame(name)
            address1=pd.DataFrame(address1)
            address2=pd.DataFrame(address2)
            address3=pd.DataFrame(address3)
            address4=pd.DataFrame(address4)
            phone=pd.DataFrame(phone)
            email=pd.DataFrame(email)
            match=pd.DataFrame(match)
            origFullName=pd.DataFrame(origFullName)
    
            # Merge all together
            
    
            thatsThem=pd.DataFrame(name.merge(address1, left_index=True,
                                 right_index=True).merge(address2, left_index=True,
                                 right_index=True).merge(address3, left_index=True,
                                 right_index=True).merge(address4, left_index=True,
                                 right_index=True).merge(phone, left_index=True,
                                 right_index=True).merge(email, left_index=True,
                                 right_index=True).merge(match, left_index=True,
                                 right_index=True).merge(origFullName, left_index=True,
                                 right_index=True))
            
    
            thatsThem.columns=['Name','Address_Line','City','State','Zip','Phone Number','Email','Match','origFullName']
            thatsThem['Address'] = thatsThem['Address_Line'].map(str) + ' ' + thatsThem['City'].map(str) + ', ' + \
            thatsThem['State'].map(str) + ' ' + thatsThem['Zip'].map(str)
            thatsThem['source'] = 'TT'
            thatsThemDict[value[0]] = thatsThem 
           
        except:
            None
 
### putting it all together 


print('\n STEP 3: Create output')       
truePeopleTemp = pd.DataFrame()  
       
for key, value in fullDict.items():
    
    df = pd.DataFrame.from_dict(fullDict[key], orient='index') 
    truePeopleTemp = truePeopleTemp.append(df, ignore_index=True) 

thatsThemTemp = pd.DataFrame()  
     
for key, value in thatsThemDict.items():
    
    thatsThemTemp = thatsThemTemp.append(value, ignore_index=True) 


thatsThemFinal = thatsThemTemp.drop(['Address_Line','City','State','Zip'], axis = 1).rename( \
                                   columns={'name':'Name'})

truePeopleFinal = truePeopleTemp.rename(columns={'name': 'Name', 'address':'Address', \
                                            'match':'Match', 'phone':'Phone Number', \
                                            'age':'Age', 'diff':'Name Distance'})

dfFinal = thatsThemFinal.append(truePeopleFinal, ignore_index=True).rename(columns={'Match':'BENE_IDX'})

#df2 = pd.merge(dfBene, dfFinal, left_on='BENE_IDX', right_on='Match') 
df2 = pd.merge(dfBene2, dfFinal, how='left', left_on='FULL_NAME', right_on='origFullName') 


addyList = []
for index, row in df2.iterrows():
    if type(row['Address']) == float:
        addyList.append('None')
    else:
        addyList.append(str(row['Address']))
        
dfNewAddy = pd.DataFrame(addyList)

dfNewAddy.columns = ['New Address']
df3 = pd.concat([df2,dfNewAddy], axis = 1)

addyFlagList = []
for index, row in df3.iterrows():
    if row['recip_city'] in row['New Address'].upper():
        addyFlagList.append('City Match')
    else:
        addyFlagList.append('')

dfAddy = pd.DataFrame(addyFlagList)

dfAddy.columns = ['Addy Match']

df4 = pd.concat([df3,dfAddy], axis = 1) 


df5 = pd.merge(df4, facPhoneDf, how='left', left_on='FULL_NAME',  right_on='MATCH_NAME')

#df5 = pd.merge(facPhoneDf, df4, how='left', on='BENE_IDX') 

deadBenes.append('dummy')
deadBenesDf = pd.DataFrame(dedup(deadBenes))


deadBenesDf.columns = ['DEAD_HIC']
df6 = pd.merge(df5, deadBenesDf, how='left', left_on='HIC',  right_on='DEAD_HIC')
df6 = df6.sort_values(by=['%s' % numType, 'Addy Match', 'Name Distance'],  ascending=[1,0,0])


dfList = list(df6)
dfDict = {}
for i, ivalue in enumerate(dfList):
    tempList = []

    for index, row in df6.iterrows():
        if type(row[ivalue]) == float:
            tempList.append('None')
        else:
            tempList.append(str(row[ivalue]))
    dfDict[ivalue] = tempList
df6 = pd.DataFrame.from_dict(dfDict)

df7 = df6[['HIC','recip_ssn','recip_id','RECIP_NAME_FIRST','RECIP_NAME_MIDDLE_INIT',\
         'RECIP_NAME_LAST','FULL_NAME','DEAD_HIC','Name','ADDRESS','recip_city','recip_state','RECIP_ZIP_CD',\
         'ZIP5','Address','Addy Match','Age','Email','TELEPHONE','Phone Number','Facility Phone',\
         'source']]

df8 = df7.rename(columns={'RECIP_NAME_LAST':'Recipient Last Name',\
                    'RECIP_NAME_FIRST':'Recipient First Name',\
                    'RECIP_NAME_MIDDLE_INIT':'Recipient Middle Initial',\
                    'ADDRESS':'SAS Address','recip_city':'Recipient City',\
                    'recip_id':'Recipient ID','RECIP_ZIP_CD':'SAS Full Zip Code',\
                    'recip_ssn':'Recipient SSN','recip_state':'Recipient State',\
                    'ZIP5':'SAS Zip Code','TELEPHONE':'SAS Telephone',\
                    'FULL_NAME':'Full Name','Address':'Web Scraped Address',\
                    'Age':'Web Scraped Age','Email':'Web Scraped Email',\
                    'Name':'Web Scraped Name','Phone Number':'Web Scraped Telephone',\
                    'source':'Web Scrape Source','Addy Match':'Address Match Indicator',\
                    'Facility Phone':'Facility Telephone','DEAD_HIC':'Deceased HIC Ind'})
    
print('\n STEP 4: Exporting the file to the same directory as User input')   

directory = file[:file.rfind('\\')]
writer = pd.ExcelWriter('%s/Bene Phone List.xlsx' % directory )
df8.to_excel(writer,'Sheet1')
writer.save()


print('\n STEP 5: Formatting File')    

def writerFormat(DIR, inFile, outFile='Same'):
    if outFile == 'Same':
        outFile = inFile
    # import pandas
    import pandas as pd
    
    # create writer
    
    # outFile format check
    
    if outFile[-5:] == '.xlsx':
        outFile = outFile[:-5] 
    elif outFile[-4:] == '.xls':
        outFile = outFile[:-4] 
    elif '.' in outFile:
        outFile = outFile[:outFile.find('.')]
        print("User Warning: Program removed 'outFile' type and replaced with '.xlsx'. \
              Sorry for your inconvenience"  )   
        
    writer = pd.ExcelWriter('%s/%s.xlsx' % (DIR, outFile), engine='xlsxwriter')

    # create xls object that contains all sheets of Excel file
    if '.' not in inFile:
        try:
            xls = pd.ExcelFile('%s/%s.xlsx' % (DIR, inFile))
            print("Program Complete. Bene Phone List file can be found in \n same folder as original HIC/SSN list.")
        except FileNotFoundError:
            try:
                xls = pd.ExcelFile('%s/%s.xls' % (DIR, inFile))
                print("The 'inFile' worked using an xls format")

            except FileNotFoundError:
                print("The 'inFile' type seems not to be compatible or non-existent.         \
                      Please make sure the 'inFile' file exist and/ or make sure             \
                      it is an xls or xlsx file and try again")
            else:
                xls = pd.ExcelFile('%s/%s.xls' % (DIR, inFile))
                sufix = '.xls' 
        else:
            xls = pd.ExcelFile('%s/%s.xlsx' % (DIR, inFile)) 
            sufix = '.xlsx' 


        
    elif '.' in inFile:
        tempSufix = inFile[-4:]
        inFile = inFile[:inFile.find('.')]
        try:
            xls = pd.ExcelFile('%s/%s.xlsx' % (DIR, inFile))
            if tempSufix != 'xlsx':
                print("The 'inFile' worked, but only after using an xlsx format")
        except FileNotFoundError:
            try:
                xls = pd.ExcelFile('%s/%s.xls' % (DIR, inFile))
                if tempSufix != '.xls':
                    print("The 'inFile' worked, but only after using an xls format")
            except FileNotFoundError:
                print("The 'inFile' type seems not to be compatible or non-existent.         \
                      Please make sure the 'inFile' file exist and/ or make sure             \
                      it is an xls or xlsx file and try again")
            else:
                xls = pd.ExcelFile('%s/%s.xls' % (DIR, inFile))
                sufix = '.xls' 
        else:
            xls = pd.ExcelFile('%s/%s.xlsx' % (DIR, inFile))
            sufix = '.xlsx' 

     
    # get sheet names of workbook to use
    worksheets = xls.sheet_names
    
#    print(worksheets)
    
    # iterate through all sheets
    for i, sheet in enumerate(worksheets):
      
        # create dataframe from excel file
        df = pd.read_excel('%s/%s%s' % (DIR, inFile, sufix), sheetname=sheet)
        
        valLen = []
        for i in df:
            valLen.append(int(df[i].map(lambda x: len(str(x))).max()))
        
        origCol = []
        newCol = []
        colCount = 0
        for col in df:
            colCount += 1
            origCol.append(str(col)) 
            newCol.append('col'+str(colCount))
        
        colLen = []
        for i in range(len(valLen)):
            colLen.append(max(valLen[i],len(origCol[i])))
           
        # write file to outDIR renamed
        df.to_excel(writer, sheet_name=sheet, startrow = 1, header = False, index = False)
        
        workbook = writer.book
           
        worksheet = writer.sheets[sheet]
        
        worksheet.autofilter(0, 0, df.shape[0], df.shape[1])
        
        
        for i in range(len(colLen)):
            worksheet.set_column(i, i, colLen[i])
            
        hex = '#%02x%02x%02x' % (131, 36, 52)
        
        header_format = workbook.add_format({
            'bold': True,
            'font_color' : '#ffffff',
            'text_wrap': False,
            'valign': 'center',
            'align' : 'center',
            'fg_color': hex,
            'border': 0})
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
         
            
    writer.close()    
        
writerFormat(directory, 'Bene Phone List')

print("The file named 'Bene Phone List' can be found in '%s'" % directory)

input("press 'Enter' to exit: ")
