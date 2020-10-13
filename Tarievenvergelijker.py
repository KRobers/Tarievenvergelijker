from selenium import webdriver
from CONFIG import *
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import os
import time
import xlsxwriter

#defines path location
path = 'C:\\Users\\kajro\\Documents\\Innova\\Pythonscripts\\Tarievenvergelijker\\'

#Path to Excel file
tarievenPath = path + 'tarieven.xlsx'
writerTarieven = pd.ExcelWriter(tarievenPath, engine='xlsxwriter')

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
except:
    pass
#Searches for the cookiebox
try:
    clickLabel = driver.find_element_by_id("CybotCookiebotDialogBodyLevelButtonAccept")
    if clickLabel.is_displayed(): #if displayed click the box
        clickLabel.click()
except:
    pass
try:
    openOptions = driver.find_element_by_css_selector(".c-gl-compare-widget__usage-link.c-link.c-link--cta.js-compare-widget-trigger.custom")
    openOptions.click()
except:
    pass

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

time.sleep(5)

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


time.sleep(2)

"""

Enkele meter 

"""

#Het eerste element
EU1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[1]/div[1]/div/div[4]/a')
EU1Jaar = EU1JaarElement.text
print(EU1Jaar)

EU1JaarEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
EU1JaarTerug = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
EU1JaarGasEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
EU1JaarVast = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text
print(EU1JaarEnkel, EU1JaarTerug, EU1JaarGasEnkel, EU1JaarVast)

#Refactoring
EU1JaarVast = EU1JaarVast[1:] #replaces the €
EU1JaarVast = float(EU1JaarVast.replace(',','.'))

EU1JaarVastStroom = EU1JaarVast / 2
EU1JaarVastGas = EU1JaarVast / 2


#Het tweede element
NED1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a')
NED1Jaar = NED1JaarElement.text
print(NED1Jaar)

NED1JaarEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
NED1JaarTerug = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
NED1JaarGasEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text
NED1JaarVast = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[4]/div').text
print(NED1JaarEnkel, NED1JaarTerug, NED1JaarGasEnkel, NED1JaarVast)

#Refactoring
NED1JaarVast = NED1JaarVast[1:] #replaces the €
NED1JaarVast = float(NED1JaarVast.replace(',','.'))

NED1JaarVastStroom = NED1JaarVast / 2
NED1JaarVastGas = NED1JaarVast / 2


#Het derde element
EU3JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a')
EU3Jaar = EU3JaarElement.text
print(EU3Jaar)

EU3JaarEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
EU3JaarTerug = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
EU3JaarGasEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text
EU3JaarVast = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[4]/div').text
print(EU3JaarEnkel, EU3JaarTerug, EU3JaarGasEnkel, EU3JaarVast)

#Refactoring
EU3JaarVast = EU3JaarVast[1:] #replaces the €
EU3JaarVast = float(EU3JaarVast.replace(',','.'))

EU3JaarVastStroom = EU3JaarVast / 2
EU3JaarVastGas = EU3JaarVast / 2

"""

Dubbele meter 

"""

try:
    wijzig = driver.find_element_by_xpath("//*[contains(text(), 'Wijzig')]")
    if wijzig.is_displayed():
        wijzig.click()
except:
    pass

#Checkbox Dubbele meter
dubbeleMeterCheckbox = driver.find_element_by_xpath('//*[@id="doublemeter1"]')
driver.execute_script("arguments[0].click();", dubbeleMeterCheckbox)

#Invullen gegevens
gaslichtNormaal = driver.find_element_by_id('usageElectricityHigh')
gaslichtNormaal.clear()
gaslichtNormaal.send_keys(verbruikNormaalTarief)

gasLichtDal = driver.find_element_by_id('usageElectricityLow')
gasLichtDal.clear()
gasLichtDal.send_keys(verbruikDalTarief)


#Vergelijk Prijs knop
vergelijkPrijs = driver.find_element_by_xpath("/html/body/section/section[1]/div[2]/div/div/aside/div[1]/div/div/div[2]/form/div/div[2]/button")
vergelijkPrijs.click()

time.sleep(5)

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

time.sleep(2)

#1st Element
EU1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[1]/div[1]/div/div[4]/a').text
print(EU1JaarElement)

EU1JaarNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
EU1JaarDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
EU1JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
EU1JaarTerugDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text
EU1JaarGasDubbel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[5]/div').text


print(EU1JaarNormaal,
EU1JaarDal,
EU1JaarTerugNormaal,
EU1JaarTerugDal,
EU1JaarGasDubbel)

#2nd Element
NED1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a').text
print(NED1JaarElement)

NED1JaarNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
NED1JaarDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
NED1JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text
NED1JaarTerugDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[4]/div').text
NED1JaarGasDubbel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[5]/div').text


print(NED1JaarNormaal,
NED1JaarDal,
NED1JaarTerugNormaal,
NED1JaarTerugDal,
NED1JaarGasDubbel)

EU3JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a').text
print(EU3JaarElement)

EU3JaarNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
EU3JaarDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
EU3JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text
EU3JaarTerugDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[4]/div').text
EU3JaarGasDubbel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[5]/div').text

print(EU3JaarNormaal,
EU3JaarDal,
EU3JaarTerugNormaal,
EU3JaarTerugDal,
EU3JaarGasDubbel)

#Creates the dataframe for the Excel file
tableTarieven = pd.DataFrame({'Enkel':[EU1JaarEnkel],
                              'Normaal':[EU1JaarNormaal],
                              'Dal':[EU1JaarDal],
                              'GasEnkel':[EU1JaarGasEnkel],
                              'GasDubbel':[EU1JaarGasDubbel],
                              'Vastrecht Stroom':[EU1JaarVastStroom],
                              'Vastrecht gas':[EU1JaarVastGas],
                              'Teruglevertarief':[EU1JaarTerug],
                              'Teruglevertarief Normaal':[EU1JaarTerugNormaal],
                              'Teruglevertarief Dal':[EU1JaarTerugDal]},
                             index=['Gaslicht.com', 'Pricewise.nl', 'Overstappen.nl', 'Independer.nl' ])


tableTarieven.to_excel(writerTarieven, sheet_name='EUR1Jaar', index=False,)
writerTarieven.save()
"""



naar excel:
tableTarieven.to_excel(writerTarieven, sheet_name='EUR_1Jaar', index=False)
writerTarieven.save()

"""