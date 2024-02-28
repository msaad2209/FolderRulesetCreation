from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.action_chains import ActionChains
from msedge.selenium_tools import Edge, EdgeOptions
import sys
import getpass
import time
import os
import re
import fnmatch
import openpyxl
import ctypes
import paramiko
import warnings
import shutil
import ctypes
RulesetP = input(f"Please enter y if you want to create ruleset else n (y/n): ") #delete
if RulesetP == 'y' or RulesetP == 'Y':
    ProjectNumber = input(f"Enter Forsta Project Pcode(4/5 digit study code): ") #delete
    pcode = str(ProjectNumber)
    #Get driver and its options and services
    edge_options = EdgeOptions()
    edge_options.use_chromium = True
    edge_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    edge_options.add_argument("--headless")
    #driver = Edge(executable_path='C:/Python3/msedgedriver.exe', options=edge_options)
    driver = Edge(executable_path='C:/Python3/msedgedriver.exe')

    driver.get("https://author.euro.confirmit.com/confirm/authoring/Confirmit.aspx")

    wait = WebDriverWait(driver, 10)
    username_input = wait.until(EC.visibility_of_element_located((By.ID, "username")))

    UserName = os.getlogin()
    usernameauto = input("If your username is "+UserName+", enter y, else n (y/n):")


    #Login to confirmit
    def login():
        max_attempts = 5
        attempts = 0

        while attempts < max_attempts:
            if usernameauto == "y" or usernameauto == "Y":
                username = UserName 
            else:
                username = input("Enter your username: ")
            password = getpass.getpass("Enter your password: ")    
            if verify_login(username, password):
                print("Login successful!")
                break
            else:
                print("Invalid username or password. Please try again.")
                attempts += 1

        if attempts == max_attempts:
            print("Maximum login attempts reached. Closing the program")
            ctypes.windll.user32.MessageBoxW(0, "Maximum login attempts reached. Closing the program", "Maximum Attempts", 0)
            driver.close()
            driver.quit()
            sys.exit()

    def verify_login(username, password):
        try:        
            username_input = driver.find_element(By.ID, "username")
            if username_input.get_attribute("value"):
                username_input.clear()
            username_input.send_keys(username)
            password_input = driver.find_element(By.ID, "password")
            if password_input.get_attribute("value"):
                password_input.clear()
            password_input.send_keys(password)
        except:
            print("Not able to pass the userid and password")

        login_button = driver.find_element(By.ID, "btnlogin")
        login_button.click()
        
        try:
            #print("Validating username and password")
            wait = WebDriverWait(driver, 5)
            username_input = wait.until(EC.visibility_of_element_located((By.ID, "__button_dataprocessing_inner")))
            return True
        except:
            return False			

    login()

    #Click DataProcessing button    
    dbutton = driver.find_element(By.ID, "__button_dataprocessing_inner")
    dbutton.click()
    RulePath = Mainfolderpath + new_folder_name + '/B. DP/b. scripts/DataMapLayout/Rules'
    xlsxfile = RulePath + pcode + "_Excel_SurveyData_sFTP.xml"
    SPSSfile = RulePath + pcode + "_SPSS_SurveyData_sFTP.xml"
    ASCIIfile = RulePath + pcode + "_ASCII_SurveyData_sFTP.xml"

    # List of file paths to be uploaded
    file_paths = [xlsxfile, SPSSfile, ASCIIfile]

    for filepath in file_paths:
        uploaded_file_name = os.path.splitext(os.path.basename(filepath))[0]
        print("Importing " + uploaded_file_name)
        #Clicking Import button 
        wait = WebDriverWait(driver, 5)
        username_input = wait.until(EC.visibility_of_element_located((By.ID, "__button_dp_rulelist")))

        rslbutton = driver.find_element(By.ID, "__button_dp_rulelist")
        rslbutton.click()

        iframe = driver.find_element(By.ID, "main_frame")
        driver.switch_to.frame(iframe)

        wait = WebDriverWait(driver, 10)
        username_input = wait.until(EC.visibility_of_element_located((By.ID, "ctl03_miImport")))

        rslbutton = driver.find_element(By.ID, "ctl03_miImport")
        rslbutton.click()

        driver.switch_to.default_content()

        #Uploading the XML file
        try:
            wait = WebDriverWait(driver, 5)
            iframemain = driver.find_element(By.ID, "main_frame")
            driver.switch_to.frame(iframemain)
            wait = WebDriverWait(driver, 5)
            username_input = wait.until(EC.visibility_of_element_located((By.ID, "jobFile")))
            file_input = driver.find_element(By.ID, "jobFile")
            file_input.send_keys(filepath)
            text_box = driver.find_element(By.ID, "panelruleName")
            text_box.clear()
            text_box.send_keys(uploaded_file_name)
            ok1 = driver.find_element(By.ID, "okbutton")
            ok1.click()
            driver.switch_to.default_content() 
        except Exception as e:
            print("Rule not imported")
            driver.switch_to.default_content()
            driver.close()
            driver.quit()
            sys.exit()

    #Creating ruleset
    print("Creating ruleset")
    wait = WebDriverWait(driver, 5)
    username_input = wait.until(EC.visibility_of_element_located((By.ID, "__button_dp_rulesetlist")))

    rslbutton = driver.find_element(By.ID, "__button_dp_rulesetlist")
    rslbutton.click()

    iframem2 = driver.find_element(By.ID, "main_frame")
    driver.switch_to.frame(iframem2)

    newruleset = driver.find_element(By.ID, "ctl02_miNew")
    newruleset.click()
    driver.switch_to.default_content()

    try:
        RulesetName = pcode + "_SurveyData(Excel+SPSS)_sFTP"

        wait = WebDriverWait(driver, 5)
        iframemain = driver.find_element(By.ID, "main_frame")
        driver.switch_to.frame(iframemain)

        text_box = driver.find_element(By.ID, "textRuleSetName")
        text_box.clear()
        text_box.send_keys(RulesetName)
        ok2 = driver.find_element(By.ID, "buttonMenu_buttonOk")
        ok2.click()
    except Exception as e:
        print("Ruleset name should be unique, delete existing ruleset with name "+RulesetName)
        driver.close()
        driver.quit()
        sys.exit()

    driver.switch_to.default_content()

    wait = WebDriverWait(driver, 5)
    iframemain = driver.find_element(By.ID, "main_frame")
    driver.switch_to.frame(iframemain)

    rulesetid = driver.find_element(By.ID, "ucGeneral_txtID")
    rulesetid_text = rulesetid.get_attribute("value")

    element = driver.find_element(By.XPATH, '//*[@id="tsRuleSetEditor"]/tbody/tr[2]/td[4]')
    element.click()

    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 5)
    iframemain = driver.find_element(By.ID, "main_frame")
    driver.switch_to.frame(iframemain)

    time.sleep(5)
    rulename1 = pcode + "_Excel_SurveyData_sFTP"
    text_box2 = driver.find_element(By.ID, "ucManagement_ruleManagementMenu_miGen_txtRuleName")
    text_box2.clear()

    actions = ActionChains(driver)

    for char in rulename1:
        text_box2.send_keys(char)
        time.sleep(0.1)

    actions.perform()

    #desired_option_text = "7701_Excel_SurveyData_sFTP"
    option_xpath = f'//div[@class="yui-ac-bd"]/ul/li[contains(text(), "{rulename1}")]'

    option_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, option_xpath)))

    option_element.click()


    ok3 = driver.find_element(By.ID, "ucManagement_ruleManagementMenu_miAdd")
    ok3.click()

    time.sleep(3)
    rulename2 = pcode + "_SPSS_SurveyData_sFTP"
    text_box2 = driver.find_element(By.ID, "ucManagement_ruleManagementMenu_miGen_txtRuleName")
    text_box2.clear()

    actions = ActionChains(driver)

    for char in rulename2:
        text_box2.send_keys(char)
        time.sleep(0.1)

    actions.perform()

    option_xpath = f'//div[@class="yui-ac-bd"]/ul/li[contains(text(), "{rulename2}")]'

    option_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, option_xpath)))

    option_element.click()

    ok3 = driver.find_element(By.ID, "ucManagement_ruleManagementMenu_miAdd")
    ok3.click()
    driver.close()
    driver.quit()
    sys.exit()

else:
    print("Folder created and rule not imoported and ruleset not created")
    driver.close()
    driver.quit()
    sys.exit()
