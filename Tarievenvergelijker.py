from selenium import webdriver
from CONFIG import *
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import os
import time
import xlsxwriter


#defines path location
path = 'C:\\Users\\kajro\\Documents\\Innova\\Pythonscripts\\Tarievenvergelijker\\'

#Chromedriver:
chromedriverPath = path + 'chromedriver.exe'
driver = webdriver.Chrome(chromedriverPath)
driver.get('https://www.gaslicht.com/energievergelijker')
driver.maximize_window()
actions = ActionChains(driver)
time.sleep(3)

#Searches for the cookiebox
try:
    cookiesBox = driver.find_element_by_id("CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll")
    if cookiesBox.is_displayed(): #if displayed click the box
        cookiesBox.click()
except: #if not found exception is made
    pass

#Searches for the cookiebox
try:
    clickLabel = driver.find_element_by_id("CybotCookiebotDialogBodyLevelButtonAccept")
    if clickLabel.is_displayed(): #if displayed click the box
        clickLabel.click()
except: #If not found exeption is made
    pass
try:
    openOptions = driver.find_element_by_css_selector(".c-gl-compare-widget__usage-link.c-link.c-link--cta.js-compare-widget-trigger.custom")
    openOptions.click()
except:#if not found exeption is made
    pass

#Path to Excel file
tarievenPath = path + 'tarieven.xlsx'

#Creates the dataframe for the Excel file
tableTarieven = pd.DataFrame(columns=['Vergelijker', 'Enkel', 'Normaal', 'Dal', 'Gas', 'Vastrecht Stroom', 'Vastrecht gas', 'Teruglevertarief', 'Teruglevertarief Normaal', 'Teruglevertarief Dal'])
writerTarieven = pd.ExcelWriter(tarievenPath, engine='xlsxwriter')

#Adds the different sites to the column
tableTarieven = tableTarieven.append({'Vergelijker': 'Gaslicht.com'}, ignore_index=True)
tableTarieven = tableTarieven.append({'Vergelijker': 'Pricewise.nl'}, ignore_index=True)
tableTarieven = tableTarieven.append({'Vergelijker': 'Overstappen.nl'}, ignore_index=True)
tableTarieven = tableTarieven.append({'Vergelijker': 'Independer.nl'}, ignore_index=True)

#Redefines variables for different sheets
tableEUR1Jaar = tableTarieven
tableEUR3Jaar = tableTarieven
tableNED1Jaar = tableTarieven
tableModel = tableTarieven

#Checkbox Innova
innovaCheckbox = driver.find_element_by_xpath('//*[@id="aanbiedersinnova-energie"]')
driver.execute_script("arguments[0].click();", innovaCheckbox)

#"Ik wil mijn verbruik zelf invullen" klikken
try:
    expandUsage = driver.find_element_by_xpath("//*[contains(text(), 'Ik wil mijn verbruik zelf invullen')]")
    if expandUsage.is_displayed():
        expandUsage.click()
except:
    pass

#Postcode
gasLichtPostal = driver.find_element_by_id('postal').send_keys(postcode)
gasLichtNmbr = driver.find_element_by_id('housenr').send_keys(huisNr)

#Verbruik niet slimme meter

gasLichtEnkel = driver.find_element_by_id('usageElectricitySingle')
gasLichtEnkel.clear()
gasLichtEnkel.send_keys(verbruikStroom)

gaslichtGas = driver.find_element_by_id('usageGas')
gaslichtGas.clear()
gaslichtGas.send_keys(verbruikGas)

gasLichtTerug = driver.find_element_by_id('terugstroomhoog')
gasLichtTerug.clear()
gasLichtTerug.send_keys(terugLevering)











#naar excel:
#tableTarieven.to_excel(writerTarieven, sheet_name='EUR_1Jaar', index=False)
#writerTarieven.save()