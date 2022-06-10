from openpyxl.descriptors.base import String
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.support import expected_conditions as EC, wait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.styles import PatternFill
import os
import openpyxl
from openpyxl.cell import Cell
import concurrent.futures
from threadify import Threadify
import threading
import multiprocessing as mp
from nameparser import HumanName
import smtplib
import math
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.cell import Cell
from time import sleep, time
import numpy as np
from array import array
import AMD
from AMD import Logins
import json
import ctypes
import time
import subprocess
import AutomationFunctions
from AutomationFunctions import Login, WaitForFunctionToWork

driver = webdriver.Chrome(ChromeDriverManager().install())

delay = 30
IsWaitingForPopup = False
IsEnteringClaimLine = False

global EnterClaimsThread
global CheckForWindowsThread 
global CheckForAleartsThread

#This function allows me to know when the page has loaded by continuously checking for an element
def WaitForElement(driver, delay, waitFor: EC, by: By, locator: str, IsWaitingForEver=False):
    #This function just makes it easier to manage
    def FindElemntWithWait():
        nonlocal driver
        nonlocal delay
        nonlocal waitFor
        nonlocal by
        nonlocal locator
        return WebDriverWait(driver, delay).until(waitFor((by, locator)))
    if IsWaitingForEver:
        return WaitForFunctionToWork(driver.find_element, 0, by, locator)
    else:
        return FindElemntWithWait()
#This function allows a thread to only effect itself while sleeping
def SleepLocalHelper(Seconds):
    sleeper = AutomationFunctions.SleepLocal(Seconds)
    while sleeper != True:
        pass
#So in Medical Billing a procedure code tells what procedure was done on a certain visit and a modifier allows doctors to sligtly alter the procedure to meet the needs of the patient. The way a modifier is presented in the Excel Spreedsheet is by combining the Procedure Code and Modifier: Ex The procedure code 97110 with an AG modifier would be written as: 97110[AG]. So this function basicly extracts the Procedure Code and checks for a Modifier, if it finds one it returns and array in the formate: [procedureCode, modifier], if not it returns an array: [procedureCode]. I have to do this in sort of a round about way because sometimes there is multiple modifiers: Ex 97110[AG][M1]  
def GetModifiersFromProcedureCode(ProcedureCode = ""):
    ProcedureCode = str(ProcedureCode).replace(" ", "")
    if "[" not in ProcedureCode:
      return None
    else:
      return ProcedureCode.replace("]", "").split("[").pop(0)
    
#This function adds the modifiers to the form on AdvancedMD by the array created in the previous function.
def AddModifersToFormByArray(ModifersArray):
    if(ModifersArray != None) and (ModifersArray != []) and (ModifersArray != ""):
        ModiferCounter = 6
        for M in ModifersArray:
            ModiferElement = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/fieldset[1]/div['+str(ModiferCounter)+']/input')
            ModiferElement.send_keys(M)
#This function clears the modifiers inputs
def ClearModifierInputs(ModifersArray):
    if(ModifersArray != None) and (ModifersArray != []) and (ModifersArray != ""):
      
        #There are multiple modifier inputs
        ModiferCounter = 6
        for M in ModifersArray:
            ModiferElement = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/fieldset[1]/div['+str(ModiferCounter)+']/input')
            ModiferElement.clear()
#In medical billing claim forms are orgonized with claim lines which signifiy specific parts of information about a specific or multiple specific visit(s)  
def AddClaimLineWithClaimForm(DateOfService, ProcedureCode, DiagnosesCode, ChargeAmount, ModifersArray, ProviderCode, UnitPriceFor97110, IsSwitchingPatients=False):

    ProviderCode = str(ProviderCode)
    UnitPriceFor97110 = str(UnitPriceFor97110)

    #Defining Element Search Critera Arrays

    ForcePaperClaimCheckboxArray = ["forcepaperclaim"]
    BegingDosDateTextBoxArray = ["begindate", "startdate"]
    EndDosDateTextBoxArray = ["enddate", "lastdate"]

    #Defining Element Varibles

    ProviderInput = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.XPATH, '//*[@id="ellProvider"]/input', True)
    DateOfServiceBeginingElement = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.ID, "txtBeginDate", True)
    DateOfServiceEndTxtBoxElement = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.ID, "txtEndDate", True)
    ProcedureCodeElement = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.XPATH, '//*[@id="ellProccode"]/input', True)
    UnitsElement = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.ID, "txtUnits", True)
    ForcePaperClaimElement = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.ID, "chkForcePaperClaim", True)

    #Filling Out Claim Form

    if IsSwitchingPatients:
       #I use a function called WaitForFunctionToWork because of internet latency, sometimes the code could not be executed due to the page loading very slowly and my code would error out. So I made a function that keeps trying the code until it works.
       WaitForFunctionToWork(ProviderInput.clear)
       WaitForFunctionToWork(ProviderInput.send_keys, ProviderCode, Keys.TAB)
    WaitForFunctionToWork(DateOfServiceBeginingElement.clear)
    WaitForFunctionToWork(DateOfServiceBeginingElement.send_keys, DateOfService, Keys.TAB)
    WaitForFunctionToWork(DateOfServiceEndTxtBoxElement.clear)
    WaitForFunctionToWork(DateOfServiceEndTxtBoxElement.send_keys, DateOfService, Keys.TAB)
    WaitForFunctionToWork(ProcedureCodeElement.clear)
    WaitForFunctionToWork(ProcedureCodeElement.send_keys, ProcedureCode, Keys.TAB)
    WaitForFunctionToWork(ClearModifierInputs, ModifersArray)
    WaitForFunctionToWork(AddModifersToFormByArray, ModifersArray)
    if "97110" in str(ProcedureCode).replace(" ", ""):
        FeeElementTxtBox = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.ID, "txtFee", True)
        WaitForFunctionToWork(FeeElementTxtBox.clear)
        WaitForFunctionToWork(FeeElementTxtBox.send_keys, UnitPriceFor97110, Keys.TAB)
        WaitForFunctionToWork(UnitsElement.clear)
        UnitsValue = round(float(ChargeAmount) / float(UnitPriceFor97110))
        WaitForFunctionToWork(UnitsElement.send_keys, str(UnitsValue), Keys.TAB)
    else:
        WaitForFunctionToWork(UnitsElement.clear)
        WaitForFunctionToWork(UnitsElement.send_keys, "1", Keys.TAB)
    i = 1
    for _ in DiagnosesCode:
        DiagnosisCodeElement = WaitForElement(driver, delay, EC.presence_of_all_elements_located, By.XPATH, '//*[@id="ellDiag10code'+str(i)+'"]/input', True)
        WaitForFunctionToWork(DiagnosisCodeElement.clear)
        WaitForFunctionToWork(DiagnosisCodeElement.send_keys, DiagnosesCode, Keys.TAB)
        i += 1

    WaitForFunctionToWork(driver.execute_script, 'document.getElementById("chkForcePaperClaim").checked = true;')
    WaitForFunctionToWork(driver.execute_script, "saveCharge()")

    #This function switches the patient in the software
def SwitchPatient(FullName, PatientDOB, IsFirstTime):
                LastName = str(FullName).split(",")[0].replace(" ", "")
                sleep(0.3)
                if IsFirstTime == False:
                    WaitForFunctionToWork(driver.execute_script, "addChargesToDB(false)")
                    try:
                        sleep(0.05)
                        driver.find_element(By.ID, "btnOK").click()
                    except:
                        pass
                window_after = driver.window_handles[1]
                driver.switch_to.window(window_after)
                driver.switch_to.frame("frmPatientInfo1")
                SearchBar = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mat-input-0"]')))
                SearchBar.clear()
                SearchBar.send_keys(FullName)
                sleep(1.5)
                try:
                    driver.find_element(By.TAG_NAME, "mat-option").click()
                except:
                    print("[Name On Spreadsheet Does Not Match Anything In Advanced MD's Database]")
                    print()
                    print("Trying To Identify The Correct Patient Name By Patients Date Of Birth]")
                    print()
                    try:
                       SearchBar.clear()
                       SearchBar.send_keys(LastName)
                       sleep(3)
                       SearchOptions = driver.find_elements(By.TAG_NAME, "mat-option")
                       print("[Patients Date Of Birth From Spread Sheet:] " + str(PatientDOB))
                       for Option in SearchOptions:
                           try:
                               SearchOptionFullName = Option.find_element(By.CSS_SELECTOR, "span > div > div.row-start-item > div > div:nth-child(1) > div").text
                               SearchOptionLastName = str(str(SearchOptionFullName).split(",")[0].replace(" ", ""))
                               print("[Trying Search Option Full Name]: " + SearchOptionFullName)
                               print("[Identified Last Name For Search Option As]: " + SearchOptionLastName)
                                                                                   
                               SearchOptionDOB = Option.find_element(By.CSS_SELECTOR, "span > div > div.row-start-item > div > div:nth-child(2) > div.additional-item.dob-item.mr-medium").text
                               print("[Trying Search Option DOB]: " + SearchOptionDOB)
                               if SearchOptionLastName == str(LastName).replace(" ", "") and SearchOptionDOB ==  str(PatientDOB).replace(" ", ""):
                                   Option.click()
                                   sleep(1.5)
                                   print("[Correct Patient Identified]")
                                   break
                               else:
                                   continue
                           except:
                               SearchOptionFullName = Option.find_element(By.CSS_SELECTOR, "span > div > div:nth-child(1) > div.name-item").text
                               SearchOptionLastName = str(str(SearchOptionFullName).split(",")[0].replace(" ", ""))
                               print("[Trying Search Option Full Name]: " + SearchOptionFullName)
                               print("[Identified Last Name For Search Option As]: " + SearchOptionLastName)
                                                                                   
                               SearchOptionDOB = Option.find_element(By.CSS_SELECTOR, "span > div > div:nth-child(2) > div:nth-child(1) > div > div.additional-item.dob-item.mr-medium").text
                               print("[Trying Search Option DOB]: " + SearchOptionDOB)
                               if SearchOptionLastName == str(LastName).replace(" ", "") and SearchOptionDOB ==  str(PatientDOB).replace(" ", ""):
                                   Option.click()
                                   sleep(1.5)
                                   print("[Correct Patient Identified]")
                                   break
                               else:
                                   continue
                    except:
                          return "Error"
                driver.switch_to.frame(0)
                driver.execute_script("tx_openChargeEntryWin()")
                sleep(1.5)
                window_after = driver.window_handles[2]
                WaitForFunctionToWork(driver.switch_to.window, window_after)
                WaitForFunctionToWork(driver.maximize_window)

def EnterClaimLine(FullName, PatientDOB, DateOfService, ProcedureCode, ChargeAmount, ModifersArray, ProviderCode, DividByAmountForUnits, Intervel, LineIsPartOfPreviouseClaim, IsFirstTime, DiagnosesCode="A99"):
            NameParser = HumanName(FullName)
            fullName = (NameParser.last + ", " + NameParser.first)
            FullName = fullName
            if LineIsPartOfPreviouseClaim == False and IsFirstTime and Intervel == 0:
                if SwitchPatient(FullName, PatientDOB, True) == "Error":
                    return "Error"
                AddClaimLineWithClaimForm(DateOfService, ProcedureCode, DiagnosesCode, ChargeAmount, ModifersArray, ProviderCode, DividByAmountForUnits, True)
                return
                
            elif LineIsPartOfPreviouseClaim and Intervel != 0:
                AddClaimLineWithClaimForm(DateOfService, ProcedureCode, DiagnosesCode, ChargeAmount, ModifersArray, ProviderCode, DividByAmountForUnits)
                return
            elif LineIsPartOfPreviouseClaim == False:
                if SwitchPatient(FullName, PatientDOB, False) == "Error":
                    return "Error"
                AddClaimLineWithClaimForm(DateOfService, ProcedureCode, DiagnosesCode, ChargeAmount, ModifersArray, ProviderCode, DividByAmountForUnits, True)
                return

def EnterClaimsFromList(username, password, officeKey, ProviderCode, PricePerUnitFor97110, PathToWorkbook, SheetName):
        book = load_workbook(filename=PathToWorkbook)
        ws = book[SheetName]
        rows = ws.iter_rows(min_row=2, min_col=1, max_col=21, values_only=True)
        PreviousFullName = ""
        PreviousDateOfService = ""
        PreviouseProcedureCode = ""
        PreviouseChargeAmount = ""
        PreviousClaimDate = ""
        PreviousClaimNumber = ""
        i = 0
        x = 0
        y = 0
        z = 0
        if __name__ == '__main__':
            CheckForWindowsThread = threading.Thread(target=AutomationFunctions.CheckForWindows, args=[driver])
            CheckForAlertsThread = threading.Thread(target=AutomationFunctions.CheckForAlerts, args=[driver])

            CheckForWindowsThread.start()
            CheckForAlertsThread.start()
        #--------THIS FUNCTION IS HERE ONLY TO HELP ORGNIZE THE CODE IN THE FOR LOOP---------------
        def AddClaimForLoop(x1, z1, i1, DateOfBirth, ProviderCode, book, ws, FileName, LineIsPartOfPreviousClaim, IsFirstTime):
                nonlocal z
                nonlocal x
                nonlocal i

                nonlocal FullName
                nonlocal DateOfService
                nonlocal ProcedureCode
                nonlocal ChargeAmount
                nonlocal ClaimDate
                nonlocal ClaimNumber

                nonlocal PreviousFullName
                nonlocal PreviousDateOfService
                nonlocal PreviouseProcedureCode
                nonlocal PreviouseChargeAmount
                nonlocal PreviousClaimDate
                nonlocal PreviousClaimNumber
                nonlocal PricePerUnitFor97110

                x = x1
                i = i1
                z = z1
                print("Entering Claim")
                if EnterClaimLine(FullName, DateOfBirth, DateOfService, ProcedureCode, ChargeAmount, ModifiersArray, ProviderCode, PricePerUnitFor97110, i, LineIsPartOfPreviousClaim, IsFirstTime) == "Error":
                    ws["A"+str(z+1)] .fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type = 'solid')
                    ws["B"+str(z+1)] .fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type = 'solid')
                    book.save(FileName)
                ws["A"+str(z+1)] .fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type = 'solid')
                ws["B"+str(z+1)] .fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type = 'solid')
                book.save(FileName)
                z = z + 1
                PreviousFullName = FullName
                PreviousDateOfService = DateOfService
                PreviouseProcedureCode = ProcedureCode
                PreviouseChargeAmount = ChargeAmount
                PreviousClaimDate = ClaimDate
                PreviousClaimNumber = ClaimNumber
        FileName = str(str(PathToWorkbook).split("/")[:-1])
        print(PathToWorkbook)
        AutomationFunctions.Login.ToAdvencedMD(driver, username, password, officeKey)
        for AccountNumber, FullName, CaseType, DateOfBirth, DateOfService, ChargeAge, ProcedureCode, Provider, ChargeAmount, InsuranceBallence, ChargeStatus, PolicySequence, Payer, SubscriberID, PlanName, PlanNumber, ClaimDate, LastBilledDate, LastPaidDate, PaidAmount, ClaimNumber in rows:
                DateOfBirth = AutomationFunctions.GetRealDate(DateOfBirth)
                DateOfService = AutomationFunctions.GetRealDate(DateOfService)
                ClaimDate = AutomationFunctions.GetRealDate(ClaimDate)
                ModifiersArray = GetModifiersFromProcedureCode(ProcedureCode)

                if str(Payer).replace(" ", "").lower() == "geico" or str(Payer).replace(" ", "").lower() == "statefarm" or str(Payer).replace(" ", "").lower() == "toblerlaw" or "usaa" in str(Payer).replace(" ", "").lower() or str(Payer).replace(" ", "").lower() == "nationwide"  or "navajo" in str(Payer).replace(" ", "").lower():
                    continue

                print("Now Doing Claim For: " + FullName)

                print("[Enter Claims Thread]: " + str("'"+PreviousFullName+"'"))
                print("[Enter Claims Thread]: " + str("'"+PreviousClaimDate+"'"))
                print("[Enter Claims Thread]: " + str("'"+DateOfService+"'"))
                print(ChargeAmount)
                print()
                y = y + 1
                if "[" in ProcedureCode:
                    ProcedureCode = str(str(ProcedureCode).split("[")[0])
                print(ProcedureCode)
                if i == 0:
                   AddClaimForLoop(1, z+1, 0, DateOfBirth, ProviderCode, book, ws, FileName, False, True)
                   i = 1
                   continue
                else:
                    if (ClaimNumber != None and ClaimNumber != ""):
                        if (ClaimNumber == PreviousClaimNumber) and (FullName == PreviousFullName):
                            AddClaimForLoop(1, z+1, 1, DateOfBirth, ProviderCode, book, ws, FileName, True, False)
                            continue
                        else:
                            AddClaimForLoop(0, z+1, 1, DateOfBirth, ProviderCode, book, ws, FileName, False, False)
                            continue
                    
                    else:
                        if (ClaimDate == PreviousClaimDate) and (FullName == PreviousFullName):
                            AddClaimForLoop(1, z+1, 1, DateOfBirth, ProviderCode, book, ws, FileName, True, False)
                            continue
                        else:
                            AddClaimForLoop(0, z+1, 1, DateOfBirth, ProviderCode, book, ws, FileName, False, False)
                            continue
        try:
            WaitForFunctionToWork(driver.execute_script, 0, "addChargesToDB(false)")
            try:
                sleep(0.05)
                driver.find_element(By.ID, "btnOK").click()
            except:
                pass
        except:
            pass
        sleep(10)
def DeleteDuplicates():
    try:
        driver.delete_all_cookies()
        AutomationFunctions.Login.ToAdvencedMD(driver, Logins.Username, Logins.Password, AMD.DrDylan.OfficeKey)
        AutomationFunctions.DeleteDuplicateClaims(driver)
    except:
        driver.quit()
        sleep(3)
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.delete_all_cookies()
        AutomationFunctions.Login.ToAdvencedMD(driver, Logins.Username, Logins.Password, AMD.DrSteve.OfficeKey)
            AutomationFunctions.DeleteDuplicateClaims(driver)
#EnterClaimsFromList(Logins.Username, Logins.Password, AMD.DrSteve.OfficeKey, AMD.DrSteve.ProviderCode, AMD.DrSteve.UnitPriceFor97110, "F:\\PrestoProfitsRCM_Automation\\PrestoProfitsRCM_Automation\\AutomationScripts\\ClaimEntry\\Global Aging 10.24.21.xlsx", "Data")
