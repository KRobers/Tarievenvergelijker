from selenium import webdriver
from CONFIG import *
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
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

def replaceKomma(value):
    value = float(value.replace(',', '.'))
    return value

def deleteEuroSign(value):
    value = value[1:]
    return value

"""
-------------------------------------
Gaslicht.com
-------------------------------------
"""

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
  prijsdetails = driver.find_elements_by_xpath("//*[contains(text(), 'Prijsdetails')]")
  for x in range(0, len(prijsdetails)):
      if prijsdetails[x].is_displayed():
          prijsdetails[x].click()
except:
    print("ERROR - PRIJSDETAILS")
    pass


time.sleep(5)

"""
-------------------------------------
Enkele meter
-------------------------------------
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
-------------------------------------
Dubbele of slimme meter
-------------------------------------
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
  prijsdetails = driver.find_elements_by_xpath("//*[contains(text(), 'Prijsdetails')]")
  for x in range(0, len(prijsdetails)):
      if prijsdetails[x].is_displayed():
          prijsdetails[x].click()
except:
    print("ERROR - PRIJSDETAILS")
    pass

time.sleep(2)

#1st Element
glEU1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[1]/div[1]/div/div[4]/a').text
print(glEU1JaarElement)#Tijdelijk

glEU1JaarNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
glEU1JaarDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
glEU1JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
glEU1JaarTerugDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text
glEU1JaarGasDubbel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[5]/div').text

#Tijdelijk
print(glEU1JaarNormaal,
glEU1JaarDal,
glEU1JaarTerugNormaal,
glEU1JaarTerugDal,
glEU1JaarGasDubbel)

#2nd Element
NED1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a').text
print(NED1JaarElement)#Tijdelijk

NED1JaarNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
NED1JaarDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
NED1JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text
NED1JaarTerugDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[4]/div').text
NED1JaarGasDubbel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[5]/div').text

#Tijdelijk
print(NED1JaarNormaal,
NED1JaarDal,
NED1JaarTerugNormaal,
NED1JaarTerugDal,
NED1JaarGasDubbel)

EU3JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a').text
print(EU3JaarElement)#Tijdelijk

EU3JaarNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
EU3JaarDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
EU3JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text
EU3JaarTerugDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[4]/div').text
EU3JaarGasDubbel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[5]/div').text

#Tijdelijk
print(EU3JaarNormaal,
EU3JaarDal,
EU3JaarTerugNormaal,
EU3JaarTerugDal,
EU3JaarGasDubbel)

"""
-------------------------------------
Modelcontract   
-------------------------------------
"""

try:
    wijzig = driver.find_element_by_xpath("//*[contains(text(), 'Wijzig')]")
    if wijzig.is_displayed():
        wijzig.click()
except:
    pass

time.sleep(1)

driver.find_element_by_xpath('//*[@id="terugstroomhoogdubbel"]').clear()
driver.find_element_by_xpath('//*[@id="terugstroomlaag"]').clear()

#Vergelijk Prijs knop
vergelijkPrijs = driver.find_element_by_xpath("/html/body/section/section[1]/div[2]/div/div/aside/div[1]/div/div/div[2]/form/div/div[2]/button")
vergelijkPrijs.click()

time.sleep(2)

glOnbepaald = driver.find_element_by_xpath('//*[@id="looptijd99"]')
driver.execute_script("arguments[0].click();", glOnbepaald)

time.sleep(2)

try:
  prijsdetails = driver.find_elements_by_xpath("//*[contains(text(), 'Prijsdetails')]")
  for x in range(0, len(prijsdetails)):
      if prijsdetails[x].is_displayed():
          prijsdetails[x].click()
except:
    print("ERROR - PRIJSDETAILS")
    pass

time.sleep(3)

glModelELement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol/li/div[1]/div[1]/div/div[4]/a').text
print(glModelELement)

glModelNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
glModelDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
glModelGas = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
glModelLever = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text

#Enkele meter

time.sleep(2)

try:
    wijzig = driver.find_element_by_xpath("//*[contains(text(), 'Wijzig')]")
    if wijzig.is_displayed():
        wijzig.click()
except:
    pass

time.sleep(1)

glSlim = driver.find_element_by_xpath('//*[@id="doublemeter1"]')
driver.execute_script("arguments[0].click();", glSlim)

vergelijkPrijs = driver.find_element_by_xpath("/html/body/section/section[1]/div[2]/div/div/aside/div[1]/div/div/div[2]/form/div/div[2]/button")
vergelijkPrijs.click()

time.sleep(3)

try:
  prijsdetails = driver.find_elements_by_xpath("//*[contains(text(), 'Prijsdetails')]")
  for x in range(0, len(prijsdetails)):
      if prijsdetails[x].is_displayed():
          prijsdetails[x].click()
except:
    print("ERROR - PRIJSDETAILS")
    pass

time.sleep(2)

glModelEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text

print(glModelEnkel, glModelNormaal, glModelDal, glModelGas, glModelLever)

"""
-------------------------------------
Overstappen.nl 
-------------------------------------
"""

driver.get('https://www.overstappen.nl/energie/vergelijken/')

time.sleep(3)

#Eerste scherm
driver.find_element_by_xpath('//*[@id="esos-widget"]/div/div/form/div/div/div[1]/div/div[1]/div[1]/div/div[1]/input').send_keys(postcode)
driver.find_element_by_xpath('//*[@id="esos-widget"]/div/div/form/div/div/div[1]/div/div[1]/div[2]/div/div[1]/div/div[1]/div/input').send_keys(huisNr)
button = driver.find_element_by_xpath('//*[@id="esos-widget"]/div/div/form/div/div/div[3]/div/button')
time.sleep(3)
button.click()
time.sleep(2)
try:
    button.click()
except:
    pass

time.sleep(3)
#Tweede scherm
select = Select(driver.find_element_by_name('currentprovider'))
select.select_by_visible_text('Weet ik niet/niet van toepassing')

button = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[4]/div[1]/div/button[1]')
button.click()

button = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[6]/div/div/button[1]')
button.click()

ovSPCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[6]/div[4]/div/div/div/div')
driver.execute_script("arguments[0].click();", ovSPCheckbox)

#---------------
#Invullen gegevens
#---------------

#Enkel
ovEnkelinput = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[6]/div[2]/div/div[1]/div/div[1]/input')
ovEnkelinput.clear()
ovEnkelinput.send_keys(verbruikStroom)
#Gas
ovGasinput = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[6]/div[2]/div/div[2]/div/div[1]/input')
ovGasinput.clear()
ovGasinput.send_keys(verbruikGas)
#Teruglever
ovTeruginput = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[6]/div[5]/div/div/div/div/div[1]/input')
ovTeruginput.clear()
ovTeruginput.send_keys(terugLevering)
#CookieButton
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/img').click()
time.sleep(3)
#Knop vergelijk
button = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div[2]/div/div/form/div[8]/div/button')
button.click()

time.sleep(5)

try:
    ovAanbieder = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/div/div[1]')
    if ovAanbieder.is_displayed():
        ovAanbieder.click()
except:
    pass

ovInnovaCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/div[2]/div/div[8]/div/input')
driver.execute_script("arguments[0].click();", ovInnovaCheckbox)

time.sleep(2)

#Meer informatie
driver.find_element_by_xpath("//*[contains(text(), 'Alle info')]").click()
time.sleep(1)

#Tarieven
driver.find_element_by_xpath("//*[contains(text(), 'Tarieven & kosten')]").click()

time.sleep(5)
"""
--------------
Gegevens
--------------
"""

try:
  all_elements = driver.find_elements_by_class_name('sc-5mx0mx-2 fZtHaY')
  for x in range(0, len(all_elements)):
      if prijsdetails[x].is_displayed():
          prijsdetails[x].click()
except:
    print("ERROR - PRIJSDETAILS")
    pass

all_elements = driver.find_elements_by_class_name('sc-5mx0mx-2 fZtHaY').text
print(all_elements)

ovEnkel, ovGas, ovVastStroom, ovVastGas = [all_elements[j] for j in (0, 1, 2, 3)]
ovTerug = '-'

print(ovEnkel, ovGas, ovVastStroom, ovVastGas)

"""
--------------
Wijziging Gegevens
--------------
"""

wijzigButton = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[1]/div[2]/button')
wijzigButton.click()

ovDubbeleCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div[2]/div[2]/div[3]/form/div[5]/div/div/input')
driver.execute_script("arguments[0].click();", ovDubbeleCheckbox)

wijzigButton = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div[2]/div[2]/div[3]/form/div[10]/div/button')
wijzigButton.click()

time.sleep(2)

ovInnovaCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/div[2]/div/div[8]/div/input')
driver.execute_script("arguments[0].click();", ovInnovaCheckbox)

time.sleep(2)

#Meer informatie
driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/button').click()
driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[2]/button').click()

time.sleep(1)

#Tarieven
driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div/div/div[1]').click()
driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[2]/div[1]/div/div[2]/div/div/div[1]').click()

time.sleep(2)

ovNormaal = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[2]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[3]/span[2]').text
ovDal = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[2]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[2]/span[2]').text
ovTerugNormaal = '-'
ovTerugDal = '-'

print(ovEnkel, ovNormaal, ovDal, ovGas, ovVastStroom, ovVastGas, ovTerug, ovTerugNormaal, ovTerugDal)

"""
-------------------------------------
Independer.nl
-------------------------------------
"""

driver.get('https://www.independer.nl/energie/intro.aspx')

time.sleep(3)

driver.find_element_by_xpath('//*[@id="cookieBar"]/div/div[3]/button').click()

#Invullen Postcode en Huisnummer
driver.find_element_by_xpath('//*[@id="salesboxForm"]/div/div/div[1]/div/div[2]/div/input').send_keys(postcode)
driver.find_element_by_xpath('//*[@id="salesboxForm"]/div/div/div[2]/div/div[2]/div/input[1]').send_keys(huisNr)
button = driver.find_element_by_xpath('//*[@id="salesboxSubmitButton"]')
button.click()

time.sleep(2)
#Volgende scherm
inCheckbox = driver.find_element_by_xpath('//*[@id="radio-1"]')
driver.execute_script("arguments[0].click();", inCheckbox)

inSPCheckbox = driver.find_element_by_xpath('//*[@id="switch-1"]')
driver.execute_script("arguments[0].click();", inSPCheckbox)

time.sleep(2)

driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-invoer-shell/ind-sidebar-layout/div/section/content/form/ind-input-group/div/div[2]/div[1]/ind-input-item/div[2]/input-item-content/ind-integer-input/div/input').send_keys(verbruikStroom)
driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-invoer-shell/ind-sidebar-layout/div/section/content/form/ind-input-group/div/ind-input-item[2]/div[2]/input-item-content/ind-integer-input/div/input').send_keys(verbruikGas)
driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-invoer-shell/ind-sidebar-layout/div/section/content/form/ind-input-group/div/ind-input-item[4]/div[2]/input-item-content/ind-integer-input/div/input').send_keys(terugLevering)

button = driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-invoer-shell/ind-sidebar-layout/div/section/content/form/ind-btn-container/button')
button.click()

time.sleep(2)

inCheckbox = driver.find_element_by_xpath('//*[@id="radio-5"]')
driver.execute_script("arguments[0].click();", inCheckbox)

inCheckbox = driver.find_element_by_xpath('//*[@id="radio-6"]')
driver.execute_script("arguments[0].click();", inCheckbox)

inCheckbox = driver.find_element_by_xpath('//*[@id="radio-10"]')
driver.execute_script("arguments[0].click();", inCheckbox)

inCheckbox = driver.find_element_by_xpath('//*[@id="radio-10"]')
driver.execute_script("arguments[0].click();", inCheckbox)

button = driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-invoer-shell/ind-sidebar-layout/div/section/content/form/ind-btn-container/button')
button.click()

time.sleep(3)

inAllContracts = driver.find_element_by_xpath('//*[@id="switch-5"]')
driver.execute_script("arguments[0].click();", inAllContracts)

time.sleep(2)

moreInfo = driver.find_element_by_id('toggle-snelle-verdieping')
moreInfo.click()








"""
#Creates the dataframe for the Excel file
tableTarieven = pd.DataFrame(data={
                              'Enkel':[glEU1JaarEnkel],
                              'Normaal':[glEU1JaarNormaal],
                              'Dal':[glEU1JaarDal],
                              'GasEnkel':[glEU1JaarGasEnkel],
                              'GasDubbel':[glEU1JaarGasDubbel],
                              'Vastrecht Stroom':[glEU1JaarVastStroom],
                              'Vastrecht gas':[glEU1JaarVastGas],
                              'Teruglevertarief':[glEU1JaarTerug],
                              'Teruglevertarief Normaal':[glEU1JaarTerugNormaal],
                              'Teruglevertarief Dal':[glEU1JaarTerugDal]},
                             index=['Gaslicht.com', 'Pricewise.nl', 'Overstappen.nl', 'Independer.nl' ])


tableTarieven.to_excel(writerTarieven, sheet_name='EUR1Jaar', index=True,)
writerTarieven.save()




naar excel:
tableTarieven.to_excel(writerTarieven, sheet_name='EUR_1Jaar', index=False)
writerTarieven.save()

"""