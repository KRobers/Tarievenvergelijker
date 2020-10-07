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



#Het eerste element
EU1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[1]/div[1]/div/div[4]/a')
EU1Jaar = EU1JaarElement.text
print(EU1Jaar)

EU1JaarEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
EU1JaarTerug = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
EU1JaarGas = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
EU1JaarVast = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text
print(EU1JaarEnkel, EU1JaarTerug, EU1JaarGas, EU1JaarVast)

EU1JaarVast = EU1JaarVast[1:] #replaces the €
EU1JaarVast = float(EU1JaarVast.replace(',','.'))

EU1JaarVastStroom = EU1JaarVast / 2
EU1JaarVastGas = EU1JaarVast / 2

tableEUR1Jaar = tableEUR1Jaar.append({'Enkel': EU1JaarEnkel}, ignore_index=True)
tableEUR1Jaar = tableEUR1Jaar.append({'Teruglevertarief': EU1JaarTerug}, ignore_index=True)
tableEUR1Jaar = tableEUR1Jaar.append({'Gas': EU1JaarGas}, ignore_index=True)
tableEUR1Jaar = tableEUR1Jaar.append({'Vastrecht Stroom': EU1JaarVastStroom}, ignore_index=True)
tableEUR1Jaar = tableEUR1Jaar.append({'Vastrecht Gas': EU1JaarVastGas}, ignore_index=True)

tableEUR1Jaar.to_excel(writerTarieven, sheet_name='EUR1Jaar', index=False, startrow=2)
writerTarieven.save()



#Het tweede element
Ned1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a')
Ned1Jaar = Ned1JaarElement.text
print(Ned1Jaar)

Ned1JaarEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
Ned1JaarTerug = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
Ned1JaarGas = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text
Ned1JaarLever = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[4]/div').text
print(Ned1JaarEnkel, Ned1JaarTerug, Ned1JaarGas, Ned1JaarLever)

#Het derde element
EU3JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a')
EU3Jaar = EU3JaarElement.text
print(EU3Jaar)

EU3JaarEnkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
EU3JaarTerug = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
EU3JaarGas = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text
EU3JaarLever = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[4]/div').text
print(EU3JaarEnkel, EU3JaarTerug, EU3JaarGas, EU3JaarLever)



#Verbruik Slimme meter






#naar excel:
#tableTarieven.to_excel(writerTarieven, sheet_name='EUR_1Jaar', index=False)
#writerTarieven.save()