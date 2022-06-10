#######################################
"""
This Python File Is a Helper File With Helper Functions Used In Many of My Automation Scripts.
I Made This Script To Better Orgonize My Code, And Preven Me From Re-coding Functions I Allready Made.
"""
#######################################
from typing import List
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
import time
import concurrent.futures
from threadify import Threadify
import threading
import multiprocessing as mp
from nameparser import HumanName
import smtplib
import math
import datetime
from openpyxl import load_workbook
from time import sleep
import numpy as np
from array import array
import json
import ctypes
import time
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import threading
import os
from ast import literal_eval

#This is a general delay I used for waiting for a webpage to load
delay = 60

#these two "Internal States" are used for Threading so the threads can pause when needed
global IsWaitingForAlert
global IsWaitingForWindow
IsWaitingForAlert = False
IsWaitingForWindow = False

#This function is used when going from one page to another, it allows me to know when AdvancedMD has actually sent me to the next page so i can more quickly execute the code i needed to
def WaitForURLToChange(driver, PreviousURL):
        CurrentURL = ""
        while PreviousURL == CurrentURL:
            CurrentURL = str(driver.current_url)

#This function is used to type into Form Inputs with a random delay between charecters to look like a human and not a bot
def TypeWithRandomDelay(driver, element, strToType, SpecailKeyCommand=None, IsDelayedBeforeAction=False, IsForDiagCode=False):
    randomStartNumber = np.random.uniform(0.03, 0.05)
    randomEndNumber = np.random.uniform(0.05, 0.11)
    print("Random Start Wait Per Char: " + str(randomStartNumber))
    print("Random End Wait Per Char: " + str(randomEndNumber))
    actions = ActionChains(driver)
    actions.move_to_element(element)
    actions.perform()
    sleep(np.random.uniform(0.01, 0.02))
    i = -1
    for Char in strToType:
        i += 1
        RandomWaitForChar = np.random.uniform(randomStartNumber, randomEndNumber)
        print("Waiting For: " + str(RandomWaitForChar) + " Seconds")
        sleep(RandomWaitForChar)
        if IsForDiagCode:
            if i == 3:
                element.send_keys(".")
        sleep(RandomWaitForChar)
        element.send_keys(Char)
    if SpecailKeyCommand != None:
        if IsDelayedBeforeAction:
            sleep(0.1)
        sleep(np.random.uniform(randomStartNumber, randomEndNumber))
        element.send_keys(SpecailKeyCommand)

#we sometimes had to get claim information from PDF's, this class helps me orgonize that information
class ClaimFromPDF:
    def __init__(self, PatientName, PatientDOB, Insurance, InsuranceID, ProcedureCodesDatesUnitsArray, DiagnossisCodes):
        self.PatientName = PatientName
        self.PatientDOB = PatientDOB
        self.Insurance = Insurance
        self.InsuranceID = InsuranceID
        self.ProcedureCodesDatesUnitsArray = ProcedureCodesDatesUnitsArray
        self.DiagnossisCodes = DiagnossisCodes


def RemoveDuplcateValuesBetweenTwoArrays(CheckArray, DeletionArray):
    for Value in CheckArray:
        try:
            DeletionArray.remove(Value)
        except:
            continue
    return DeletionArray

#This element allows me to take an element and check for an ID so in the furter i could  refrence that element again (I was sort of using my own AI to automaticly find element's on webpages based on Key-Words), so i would save the ID's of the elements my AI found to save time in the furture
def ElementHasAnId(element):
    try:
        if element.get_attribute("id") == "" or element.get_attribute("id") == None:
            return False
        else:
            return True
    except:
        return False
#This is another function that helped my Element Finding AI because incase the Element did not have an ID attrobute than I would get the XPATH of that element to save for later
def GetXPATH(driver, element):
    XpathString = ""
    NextElement = element
    DoseElementHaveAnID = ElementHasAnId(element)
    if ElementHasAnId(element):
        return '//[@id="'+str(element.get_attribute("id"))+'"]'
    while True:
        try:
            
            if ElementHasAnId(NextElement):
                XpathString += "/" + '[@id="'+str(NextElement.get_attribute("id"))+'"]'
            else:
                ParrentElement = NextElement.find_element(By.XPATH, "..")
                ParrentChildren = ParrentElement.find_elements(By.TAG_NAME, str(element.tag_name))
                i = 0
                Index = 0
                for Child in ParrentChildren:
                    if Child == element:
                        Index = i
                    else:
                        i += 1
                XpathString += "/" + NextElement.find_element(By.XPATH, "..").tag_name + "["+str(Index)+"]"
            NextElement = NextElement.find_element(By.XPATH, "..")
        except Exception as EX:
            print(EX)
            FinalString = ""
            XpathArray = XpathString.split("/")[::-1]
            i = 1
            for Section in XpathArray:
                if i == len(XpathArray):
                    FinalString += "/" + Section
                else:
                    FinalString += "//" + Section
                i += 1
            return FinalString + element.tag_name
#This function is used to more accurately and quickly see if the webpage has loaded or not
def IsElementClickable(driver, element):
    try:
        print(GetXPATH(driver, element))
        if WebDriverWait(driver, 0.0001).until(EC.element_to_be_clickable((By.XPATH, GetXPATH(driver, element)))) == None:
            return False
        else:
            return True
    except Exception as EX:
        print(EX)
        return False
#This is my sort of AI that can take in a key word and use the packge FuzzyWuzzy to search through all the elements on the webpage and see what elements are closes to the keyword and then return that element
class AutoFindElement:
    def FindElementsByRefrenceStringHelper(driver, by, value):
            if by == By.ID:
                return driver.find_elements(By.ID, value)
            elif by == By.TAG_NAME:
                return driver.find_elements(By.TAG_NAME, value)
            elif by == By.CLASS_NAME:
                return driver.find_elements(By.CLASS_NAME, value)
            elif by == By.CSS_SELECTOR:
                return driver.find_elements(By.CSS_SELECTOR, value)
            elif by == By.LINK_TEXT:
                return driver.find_elements(By.LINK_TEXT, value)
            elif by == By.NAME:
                return driver.find_elements(By.NAME, value)
            elif by == By.PARTIAL_LINK_TEXT:
                return driver.find_elements(By.PARTIAL_LINK_TEXT, value)
            elif by == By.XPATH:
                return driver.find_elements(By.XPATH, value)
            else:
                    return None
    def FindElementsByRefrenceString(driver, KeyStringValues, MatchRatio, CheckText=True):
        if CheckText:
            ListOfBySelctors = [By.ID, By.NAME, By.CLASS_NAME, By.LINK_TEXT, By.PARTIAL_LINK_TEXT]
        else:
            ListOfBySelctors = [By.ID, By.NAME, By.CLASS_NAME]
        ListOfElements = [] 
        for selctor in ListOfBySelctors:
            for KeyString in KeyStringValues:
                ElementsFoundByHelperList = FindElementsByRefrenceStringHelper(driver, selctor, KeyString)
                for Element in ElementsFoundByHelperList:
                    if Element == None or Element == "":
                        pass
                    else:
                        ListOfElements.append(Element)
        return ListOfElements
    def AutoFindElement(driver, KeyWords, CheckElementText=True):
        sleep(3)
        global ElementsDict
        global MatchPrecentages
        ElementsDict = {}
        MatchPrecentages = []
        def MainLoop():
            nonlocal driver, KeyWords, CheckElementText
            ListOfSelectors = ["id", "name", "value", "tittle", "class", "text"]
            ListOfElements = driver.find_elements(By.XPATH, "//*")
            print(ListOfElements)
            for Selctor in ListOfSelectors:
                for Element in ListOfElements:
                    if Element.is_displayed() and IsElementClickable(driver, Element):
                        for KeyWord in KeyWords:
                                if CheckElementText == False and Selctor == "text":
                                    continue
                                elif CheckElementText and Selctor == "text":
                                    try:
                                        SelctorValue = Element.text
                                    except:
                                        continue
                                try:
                                    if Selctor != "text":
                                        SelctorValue = Element.get_attribute(Selctor)
                                    print(Selctor)
                                    print(SelctorValue)
                                    StringRelationPrecentage = GetStringRelationPrecentage(SelctorValue, KeyWord)
                                    MatchRatio = GetStringRelationPrecentage(KeyWord, Element.get_attribute("value"))
                                    MatchPrecentages.append(MatchRatio)
                                    ElementsDict[MatchRatio] = Element
                                    print(str(MatchRatio) + " : " + Element.get_attribute("value"))
                                    print(ElementsDict)
                                except Exception as EX:
                                    print(EX)
                                    continue
                    else:
                        continue
        MainLoop()
        if ElementsDict[max(MatchPrecentages)] == None:
            x = 10
            while x >= 1:
                try:
                    driver.switch_to.parent_frame()
                except:
                    pass
                x = x - 1
            ListOfIframes = driver.find_elements(By.TAG_NAME, "iframe")
            try:
                for Iframe in ListOfIframes:
                        print(Iframe.get_attribute("name"))
                        driver.switch_to.frame(Iframe)
                        MainLoop()
                        Results2 = ElementsDict[max(MatchPrecentages)]
                        if ElementsDict[max(MatchPrecentages)] == None:
                            driver.switch_to.parent_frame()
                            continue
                        else:
                            return Results2
            except Exception as EX:
                     print(EX)

        else:
            return ElementsDict[max(MatchPrecentages)]
    def AndClick(driver, KeyWords, CheckElementText=True):
            ExcludeElemnts = []
            sleep(3)
            global ElementsDict
            global MatchPrecentages
            ElementsDict = {}
            MatchPrecentages = []
            def MainLoopHelper(Exclusions=[]):
                 nonlocal driver, KeyWords, CheckElementText
                 ListOfSelectors = ["id", "name", "value", "tittle", "class", "text"]
                 ListOfElements = driver.find_elements(By.XPATH, "//*")
                 ListOfElements = RemoveDuplcateValuesBetweenTwoArrays(Exclusions, ListOfElements)
                 print(ListOfElements)
                 ListOfIframes = driver.find_elements(By.TAG_NAME, "iframe")
                 for Selctor in ListOfSelectors:
                                    for Element in ListOfElements:
                                        for KeyWord in KeyWords:
                                                if CheckElementText == False and Selctor == "text":
                                                    continue
                                                elif CheckElementText and Selctor == "text":
                                                    try:
                                                        SelctorValue = Element.text
                                                    except:
                                                        continue
                                                try:
                                                    if Selctor != "text":
                                                        SelctorValue = Element.get_attribute(Selctor)
                                                    print(Selctor)
                                                    print(SelctorValue)
                                                    StringRelationPrecentage = GetStringRelationPrecentage(SelctorValue, KeyWord)
                                                    MatchRatio = GetStringRelationPrecentage(KeyWord, Element.get_attribute("value"))
                                                    MatchPrecentages.append(MatchRatio)
                                                    ElementsDict[MatchRatio] = Element
                                                    print(str(MatchRatio) + " : " + Element.get_attribute("value"))
                                                    print(ElementsDict)
                                                except Exception as EX:
                                                    print(EX)
                                                    continue
                                        else:
                                            continue
            def MainLoop(Exclusions=[]):
                nonlocal driver, KeyWords, CheckElementText
                ListOfSelectors = ["id", "name", "value", "tittle", "class", "text"]
                ListOfElements = driver.find_elements(By.XPATH, "//*")
                print(ListOfElements)
                ListOfIframes = driver.find_elements(By.TAG_NAME, "iframe")
                try:
                    if len(ListOfIframes) > 0:
                        for Iframe in ListOfIframes:
                                print(Iframe.get_attribute("name"))
                                driver.switch_to.frame(Iframe)
                                MainLoopHelper(Exclusions)
                    else:
                        MainLoopHelper(Exclusions)
                except Exception as EX:
                    print(EX)
            driver.switch_to.default_content()
            MainLoop()
            while True:
                try:
                    ElementsDict[max(MatchPrecentages)].click()
                    return
                except:
                    driver.switch_to.default_content()
                    ExcludeElemnts.append(ElementsDict[max(MatchPrecentages)])
                    MainLoop(ExcludeElemnts)
#This function saved time by allowing me to check for Alerts and popup windows before executing the code, it also prevented my code from failing due to alerts and window popups
def CheckForWindowsAndAlertsThenExecuteFunction(function, *args):
        if IsWaitingForAlert:
            while IsWaitingForAlert:
                sleep(0.1)
        elif IsWaitingForWindow:
            while IsWaitingForWindow:
                pass
        else:
            return function(*args)

def WaitForFunctionToWork(function, *args):
    FucntionSuccessful = False
    while FucntionSuccessful == False:
        if IsWaitingForAlert:
            while IsWaitingForAlert:
                sleep(0.1)
        if IsWaitingForWindow:
            while IsWaitingForWindow:
                pass
        try:
            sleep(0.2)
            return function(*args)
            FucntionSuccessful = True
        except:
            continue
#This waits for an element to be found so i can preven my code from runing while the webpage is stil loading
def WaitForStatus(driver, Element, FindBy):
        IsLoaded = False
        while IsLoaded == False:

            if FindBy.upper() == "ID":
                 try:
                        print(str(driver.find_element_by_id(str(Element)).text))
                        IsLoaded = True
                        break
                 except:
                    continue
            elif FindBy.upper() == "XPATH":
                try:
                        print(str(driver.find_element_by_xpath(str(Element)).text))
                        IsLoaded = True
                        break
                except:
                    continue
            elif FindBy.upper() == "CSS":
                try:
                        driver.find_element(By.CSS_SELECTOR,str(Element))
                        IsLoaded = True
                        break
                except:
                   continue
            elif FindBy.Upper("ELEMENT"):
                try:
                        str(Element.text)
                        IsLoaded = True
                        break
                except:
                   continue
#in AdvancedMD the webpage would load but AdvancedMD would be displaying a loading spinner which would realy mess up my code, so i made a frunction that waited for the spiner to stop, before my code executes
def WaitForLoadingToComplete(driver):
    IsLoading = True
    while IsLoading:
        try:
            if "display:none" in str(driver.find_element(By.ID, "LoadingIndicator").get_attribute("style")).replace(" ", ""):
                IsLoading = False
            else:
                IsLoading = True
        except:
            pass
#sometimes the webpage would have a loaded status but some individual elements would still be "Artificaly inactive" becuase AdvancedMD disables some elements while loading (i guess to prevent impatient useres from spamming buttons and makeing the page take longer to load) so thats why i made this funciton
def WaitForElement(driver, delay, waitFor: EC, by: By, locator: str, IsWaitingForEver=False):
    def FindElemntWithWait():
        nonlocal driver
        nonlocal delay
        nonlocal waitFor
        nonlocal by
        nonlocal locator
        return WebDriverWait(driver, 60).until(waitFor((by, locator)))
    if IsWaitingForEver:
        return WaitForFunctionToWork(driver.find_element, by, locator)
    else:
        return FindElemntWithWait()
#Sometimes AdvancedMD would enter Duplicate Charges because they had to connect to the Clients POS, Medical Notes System which caused a lot of problomes, so i had made this function to get rid of duplicate charges in our system
def DeleteDuplicateClaims(driver):
    window_after = driver.window_handles[0]
    driver.switch_to.window(window_after)
    driver.close()
    window_after = driver.window_handles[0]
    driver.switch_to.window(window_after)
    driver.maximize_window()
    driver.switch_to.parent_frame()
    driver.switch_to.parent_frame()
    driver.find_element(By.ID, "mnuClaimsCenter").click()
    while len(driver.window_handles) <= 1:
        pass 

    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    driver.maximize_window()
    sleep(3)
    try:
        alert = driver.switch_to.alert
        alert.accept()
        sleep(0.1)
    except:
        pass
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

    table = WaitForFunctionToWork(WaitForElement, driver, 60, EC.alert_is_present, By.XPATH, "//table[@id='tblRevVisit']//tbody", True)
    VisitsToDelete = []
    i = 1
    for row in table.find_elements(By.TAG_NAME, "tr"):
        try:
            if "visit" in str(row.get_attribute("id")).replace(" ", ""):
                continue
            else:
                try:
                    if "color:red" in str(row.get_attribute("style")).replace(" ", ""):
                       visitID = str(str(row.get_attribute("visitid")).replace(" ", ""))
                    else:
                        continue
                    if visitID in VisitsToDelete:
                        continue
                    else:
                        VisitsToDelete.append(str(visitID))
                except:
                    continue

        except:
            pass
    DeletedVisits = []
    DeletedVisitsForPrint = VisitsToDelete
    for visit in VisitsToDelete:
                if visit in DeletedVisits:
                    continue
                else:
                   sleep(0.3)
                   WaitForLoadingToComplete(driver)
                   row = driver.find_element(By.ID, "visit"+visit+"")
                   WaitForFunctionToWork(driver.execute_script, 'document.getElementById("chk'+visit+'").checked = true;')
                   WaitForFunctionToWork(driver.find_element,By.ID, "btnRevVoidVisit").click()
                   sleep(0.5)
                   try:
                        alert = driver.switch_to.alert
                        alert.accept()
                        sleep(0.3)
                   except:
                        pass
                   try:
                        alert = driver.switch_to.alert
                        alert.accept()
                   except:
                        pass
                   WaitForFunctionToWork(driver.execute_script, 'document.getElementById("chk'+visit+'").checked = false;')
                   WaitForFunctionToWork(DeletedVisits.append, str(visit))
                   sleep(0.3)
                   WaitForFunctionToWork(driver.find_element, By.ID, "btnRevRefresh").click()
                   WaitForLoadingToComplete(driver)
                   sleep(0.5)
                   WaitForFunctionToWork(driver.execute_script, 'document.getElementById("chk'+visit+'").checked = false;')
                   try:
                        alert = driver.switch_to.alert
                        alert.accept()
                        sleep(0.3)
                   except:
                        pass
                   try:
                        alert = driver.switch_to.alert
                        alert.accept()
                   except:
                        pass
                   DeletedVisitsForPrint.pop(0)
                   print(DeletedVisitsForPrint)
                   os.system('cls')

                   sleep(2)
#This was another function that was used by my "AI" to get the relationship ratio between the Keywords and Elements
def GetStringRelationPrecentage(Str1, Str2):
    Str1 = str(Str1).lower().replace(" ", "")
    Str2 = str(Str2).lower().replace(" ", "")
    return int(fuzz.token_sort_ratio(Str1, Str2))
def GetRealDate(Date, DateSeporatrors="-"):
    DateTimeArray = str(Date).split(" ")
    try:
        DateArray = str(DateTimeArray[0]).split("-")
        return DateArray[0] + "/" + DateArray[1] + "/" + DateArray[2]
    except:
        DateArray = str(DateTimeArray[0]).split("/")
        return DateArray[0] + "/" + DateArray[1] + "/" + DateArray[2]

#This is used in threads so when they sleep it dose not effect the main thread
def SleepLocal(Seconds):
    TimeStart = time.time()
    while  float(time.time()) - float(TimeStart) <= Seconds:
        print(float(time.time()) - float(TimeStart))
    return True
 
def CheckForAlearts(driver):
    while True:
        try:
          alert = driver.switch_to.alert
          IsWaitingForAlert = True
          print("[Checking For Alearts Thread]: Detected Aleart: " + alert.text)
          alert.accept()
          sleep(0.3)
          IsWaitingForAlert = False
        except:
            continue 
            
def CheckForWindows(driver):
    while True:
        try:
            window_after = driver.window_handles[3]
            IsWaitingForWindow = True
            driver.switch_to.window(window_after)
            driver.maximize_window()
            print("[Checking For Window Thread]: Detected Window")
            AutoFind.FindElementAuto(driver, ["ok", "accept"]).click()
            window_after = driver.window_handles[2]
            driver.switch_to.window(window_after)
            IsWaitingForWindow = False
        except:
            continue
        
#This is also used by my "AI", it gets all the label elements that match a set of keywords and then finds the "for" attrobute which contains the inputs ID to reaturn the Input element for that label element
def GetInputElementByLabelElementsOnPage(driver, KeyWordsArray):
    Label_Elements = driver.find_elements(By.TAG_NAME, "label")
    LabelElementsDict= {}
    MatchPrecentages = []
    for Label_Element in Label_Elements:
        for KeyWord in KeyWordsArray:
            MatchPrecentage = GetStringRelationPrecentage(KeyWord, Label_Element.text)
            print("'" + Label_Element.text + "'" + " KeyWord: '" + str(KeyWord) + "': " + str(MatchPrecentage))
            LabelElementsDict[MatchPrecentage] = Label_Element
            MatchPrecentages.append(MatchPrecentage)

    return driver.find_element(By.ID, str(LabelElementsDict[max(MatchPrecentages)].get_attribute("for")))

#This is the same as the previous function except insted of returning one input element, it returns an array of all the elements thatt matched within a certain precentage
def GetInputFromKeyWordsArray(driver, KeyWordsArray):
    InputElementsDict = {}
    MatchPrecentages = []
    Input_Elements = driver.find_elements(By.TAG_NAME, "input")
    for Input_Element in Input_Elements:
        if str(Input_Element.get_attribute("type")).lower().replace(" ", "") == "hidden":
            continue
        for KeyWord in KeyWordsArray:
            MatchRatio = GetStringRelationPrecentage(KeyWord, Input_Element.text)
            MatchPrecentages.append(MatchRatio)
            InputElementsDict[MatchRatio] = Input_Element
            print()
            print(str(MatchRatio) + " : " + Input_Element.text )

            #Trying Value Attribute

            MatchRatio = GetStringRelationPrecentage(KeyWord, Input_Element.get_attribute("value"))
            MatchPrecentages.append(MatchRatio)
            InputElementsDict[MatchRatio] = Input_Element
            print(str(MatchRatio) + " : " + Input_Element.get_attribute("value"))
    print()
    print(InputElementsDict)
    print()
    return InputElementsDict[max(MatchPrecentages)]
  
#these next three functions where used to get certain information from a clasified database containg our clients NPI, and certain login information
def FormatCommandStringFromSelctorDB(Value, CommandString):
    return str(CommandString).replace("%DRIVER%", "driver").replace("%VALUE%", "Value").replace("%KEY_WORDS_ARRAY%", "Value")
  
def ReturnElementFromCommandString(driver, Value, CommandString):
    FormatCommandStringFromSelctorDB(Value, CommandString)

def GetAgingSiteRowFormDB(AgingSiteName, PathToDB="E:\PrestoProfitsRCM_Automation\PrestoProfitsRCM_Automation\AgingReport\AgingReportMGR.accdb"):
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+str(PathToDB)+';')
        cursor = conn.cursor()
        ColumnNames = []
        for column in cursor.columns(table='AgingWebsitesWithoutFormInfo'):
           ColumnNames.append(column.column_name)
        cursor.execute('SELECT * FROM AgingWebsitesWithoutFormInfo')
        for row in cursor.fetchall():
            LookupSiteRowFormDB = {}
            i = 0
            for ColumnValue in row:
                      LookupSiteRowFormDB[ColumnNames[i]] = row[i]
                      i += 1
        if LookupSiteRowFormDB["LookupSiteName"].upper().replace(" ", "") == AgingSiteName.upper().replace(" ", ""):
            print(LookupSiteRowFormDB)
            return LookupSiteRowFormDB
#in that database i talked about above i hade saved python code that I wanted to execute if certain criteria where met, thats what the next two functions do
def ExecSelectorStringAndReturnElement(expression, Value, driver):
    exec(f"""locals()['temp'] = {expression}""")
    print(Value)
    return locals()['temp']
def GetAndRunStringForSelector(driver, SelectorValue, LookupStringOrArray, DbDict, PathToDB="E:\PrestoProfitsRCM_Automation\PrestoProfitsRCM_Automation\AgingReport\AgingReportMGR.accdb"):
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+str(PathToDB)+';')
        cursor = conn.cursor()
        ColumnNames = []
        DbRows = []
        for column in cursor.columns(table='SelectorLookup'):
           ColumnNames.append(column.column_name)
        cursor.execute('SELECT * FROM SelectorLookup')
        for row in cursor.fetchall():
            LookupSiteRowFromDB = {}
            i = 0
            for ColumnValue in row:
                      LookupSiteRowFromDB[ColumnNames[i]] = row[i]
                      i += 1
                      DbRows.append(LookupSiteRowFromDB)
        i = 0
        for Row in DbRows:
            if str(Row["Selector"]).replace(" ", "") == SelectorValue:
                loc = {}
                print(DbDict["IsLoginButtonKeyWordsAnArray"])
                if DbDict["IsLoginButtonKeyWordsAnArray"] == True:
                    print(str(LookupStringOrArray).split(","))
                    print(str(FormatCommandStringFromSelctorDB(str(LookupStringOrArray).split(","), str(Row["CodeToExecute"]))))
                    CodeToExecuteString = str(FormatCommandStringFromSelctorDB("Value", str(Row["CodeToExecute"])))
                    return ExecSelectorStringAndReturnElement(str(CodeToExecuteString),str(LookupStringOrArray).split(","), driver)
                else:
                    CodeToExecuteString = str(FormatCommandStringFromSelctorDB("Value", str(Row["CodeToExecute"])))
                    return ExecSelectorStringAndReturnElement(str(CodeToExecuteString),"'" + LookupStringOrArray + "'", driver)
            i = i + 1

#This function was used to delete my cache, becuase sometimes AdvancedMD would mess up if you had too much in your cache
def delete_cache(driver):
    driver.execute_script("window.open('');")
    sleep(2)
    driver.switch_to.window(driver.window_handles[-1])
    sleep(2)
    driver.get('chrome://settings/clearBrowserData') # for old chromedriver versions use cleardriverData
    sleep(2)
    actions = ActionChains(driver) 
    actions.send_keys(Keys.TAB * 3 + Keys.DOWN * 3) # send right combination
    actions.perform()
    sleep(2)
    actions = ActionChains(driver) 
    actions.send_keys(Keys.TAB * 4 + Keys.ENTER) # confirm
    actions.perform()
    sleep(5) # wait some time to finish
    driver.close() # close this tab
    driver.switch_to.window(driver.window_handles[0]) # switch back
    
#This was used to allow me to quickly login to diffrent websites in a more orgnized way
class Login:
    def ToAdvencedMD(driver, username, password, officeKey, Function=0):
        delete_cache(driver)
        driver.get("https://login.advancedmd.com/")

        driver.switch_to.frame(0)
        driver.find_element(By.NAME, "loginname").send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password)
        driver.find_element(By.NAME, "officeKey").send_keys(officeKey)
        WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn"))).click()

        while len(driver.window_handles) <= 1:
            pass

        window_after = driver.window_handles[1]
        driver.minimize_window()
        driver.switch_to.window(window_after)
        driver.maximize_window()
        if Function == 0:
            WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, 'mnuPatientInfo')))
            driver.execute_script("document.getElementById('mnuPatientInfo').click()")
            sleep(3)
            driver.switch_to.frame(0)
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cdk-drop-list-0"]/div[8]')))
            driver.execute_script('document.querySelector("#cdk-drop-list-0 > div:nth-child(9) > amds-left-panel-item > a").click()')
            sleep(3)
        elif Function == 1:
            sleep(12)
            driver.find_element(By.XPATH, '//*[@id="amds-navbar"]/ul/li[5]/a').click()
            sleep(3)
            driver.find_element(By.XPATH, '//*[@id="amds-navbar"]/ul/li[5]/div/ul/li[2]/a').click()
            sleep(3)
            driver.find_element(By.XPATH, '//*[@id="amds-navbar"]/ul/li[5]/div/ul/li[2]/div/ul/li[1]/a').click()
            while len(driver.window_handles) <= 2:
                pass
            window_after = driver.window_handles[2]
            driver.minimize_window()
            driver.switch_to.window(window_after)
            sleep(3) 
            driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div[1]/ul/li[2]/a').click()
            sleep(2)
            WaitForFunctionToWork(driver.execute_script, 'document.getElementById("chkActionDue").checked = false;')
            driver.find_element(By.XPATH, '//*[@id="divAccount"]/form/div[2]/amds-search/div/span/button').click()
            sleep(1)
            TypeWithRandomDelay(driver, driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div[2]/div[1]/div/input'), "BCBS CHECK STATUS")
            sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div[2]/div[1]/div/button').click()
            sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div[3]/div[2]/button').click()

        elif Function == 2:
            WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, 'mnuPatientInfo')))
            driver.execute_script("document.getElementById('mnuPatientInfo').click()")
            sleep(10)
            driver.switch_to.frame(0)
            driver.find_element(By.XPATH, '//*[@id="cdk-drop-list-0"]/div[14]').click()
            sleep(3)
        elif Function == 3:
            WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, 'mnuPatientInfo')))
            driver.execute_script("document.getElementById('mnuPatientInfo').click()")
            sleep(10)
            driver.switch_to.frame(0)
            driver.find_element(By.XPATH, '//*[@id="cdk-drop-list-0"]/div[4]').click()
            sleep(3)
        elif Function == 4:
            WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, 'mnuPatientInfo')))
            driver.execute_script("document.getElementById('mnuPatientInfo').click()")
            sleep(10)
            driver.switch_to.frame(0)
        elif Function == 5:
            WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, 'mnuPatientInfo')))
            driver.execute_script("document.getElementById('mnuPatientInfo').click()")
            sleep(10)
            driver.switch_to.frame(0)
            driver.find_element(By.XPATH, '//*[@id="cdk-drop-list-0"]/div[8]/amds-left-panel-item/a/i').click()

    def ToBCBSAZ(driver, Website, Username, Password):
        driver.get(Website)
        driver.maximize_window()
        UsernameInput = driver.find_element(By.ID, "lockedcontent_0_maincolumn_2_twocolumn2fb4d204091d44aa08196ef423877fd9f_0_ToolbarUsernameControl")
        PasswordInput = driver.find_element(By.ID, "lockedcontent_0_maincolumn_2_twocolumn2fb4d204091d44aa08196ef423877fd9f_0_ToolbarPasswordControl")
        LoginBtn = driver.find_element(By.ID, "lockedcontent_0_maincolumn_2_twocolumn2fb4d204091d44aa08196ef423877fd9f_0_ToolbarLoginButtonControl")
        TypeWithRandomDelay(driver, UsernameInput, str(Username))
        TypeWithRandomDelay(driver, PasswordInput, str(Password))
        sleep(np.random.uniform(0.5, 1))
        LoginBtn.click()

        #Waiting For Page To Load

        WaitForURLToChange(driver, str(driver.current_url))

        #Clicking On Search Claim Button
        
        #Delaying For Bot Detection
        sleep(np.random.uniform(9, 10))
        IsLoaded = False
        while IsLoaded == False:
            LoadingElement = driver.find_element(By.XPATH, '//*[@id="loadingContent"]')
            if "display: block;" in str(LoadingElement.get_attribute("style")):
                continue
            else:
                IsLoaded = True
        driver.find_element(By.ID, "claimsQuicksearchTab").click()

    def Aetna(driver, LookupSiteRowFormDB):
        AgingSiteDict = GetAgingSiteRowFormDB("AETNA")
        URL = AgingSiteDict["LookupSiteURL"]

        username = AgingSiteDict["Username"]
        password = AgingSiteDict["Password"]
        driver.get(URL)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.delete_all_cookies()
        sleep(5)
        GetInputElementByLabelElementsOnPage(driver, ["userid", "username"]).send_keys(str(username))
        GetInputElementByLabelElementsOnPage(driver, ["password"]).send_keys(str(password))
        GetAndRunStringForSelector(driver, AgingSiteDict["LoginButtonSelector"], AgingSiteDict["LoginButtonKeyWords"], AgingSiteDict)
        sleep(5)
        if AgingSiteDict["IsThereSomthingToClickOnForClaimStatus"]:
            CheckClaimStatusBtn = GetAndRunStringForSelector(driver, AgingSiteDict["ClaimStatusButtonSelctor"], AgingSiteDict["ClaimStatusButtonKeyWords"], AgingSiteDict)

        print(AgingSiteDict["ClaimStatusButtonSelctor"])
        print(AgingSiteDict["LoginButtonKeyWords"])
        input()
