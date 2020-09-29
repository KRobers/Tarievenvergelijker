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

vergelijkPrijs = driver.find_element_by_xpath("/html/body/section/section[1]/div[2]/div/div/aside/div[1]/div/div/div[2]/form/div/div[2]/button")
vergelijkPrijs.click()

time.sleep(3)

try:
    prijsDetails1 = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[2]/div[2]/ul/li[1]')
    if prijsDetails1.is_displayed():
        prijsDetails1.click()

    prijsDetails2 = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[2]/div[2]/ul/li[1]')
    if prijsDetails2.is_displayed():
        prijsDetails2.click()

    prijsDetails3 = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[2]/div[2]/ul/li[1]')
    if prijsDetails3.is_displayed():
        prijsDetails3.click()
except:
    pass

time.sleep(3)

#Het eerste element
naam1Element = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[1]/div[1]/div/div[4]/a')
naam1 = naam1Element.text
print(naam1)

naam1Enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
naam1Terug = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
naam1Gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
naam1Lever = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text
print(naam1Enkel, naam1Terug, naam1Gas, naam1Lever)

#Het tweede element
naam2Element = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a')
naam2 = naam2Element.text
print(naam2)

naam2Enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
naam2Terug = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
naam2Gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text
naam2Lever = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[4]/div').text
print(naam2Enkel, naam2Terug, naam2Gas, naam2Lever)

#Het derde element
naam3Element = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a')
naam3 = naam3Element.text
print(naam3)

naam3Enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
naam3Terug = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
naam3Gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text
naam3Lever = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[4]/div').text
print(naam3Enkel, naam3Terug, naam3Gas, naam3Lever)

if naam1 == 'Europese groene stroom en gas vast 1 jaar Actie':
    EUR1Jaar = naam1



#naar excel:
#tableTarieven.to_excel(writerTarieven, sheet_name='EUR_1Jaar', index=False)
#writerTarieven.save()