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

#Debugging
gaslicht = input("Gaslicht y/n: ")
overstappen = input("Overstappen y/n: ")
independer = input("Independer y/n: ")
pricewise = input("Pricewise y/n: ")

#Chromedriver:
chromedriverPath = path + 'chromedriver.exe'
driver = webdriver.Chrome(chromedriverPath)
actions = ActionChains(driver)
time.sleep(3)

def replaceKomma(value):
    value = float(value.replace(',', '.'))
    return value

def deleteEuroSign(value):
    value = value[1:]
    return value

class product:
    enkel = ""
    normaal = ""
    dal = ""
    vastrecht = ""
    teruglever = ""
    terugleverNormaal = ""
    terugleverDal = ""
    gas = ""

#Initialize objects

#Gaslicht
glEU1 = product()
glEU3 = product()
glNL1 = product()
glMOD = product()

#Pricewise
pwEU1 = product()
pwEU3 = product()
pwNL1 = product()
pwMOD = product()

#Overstappen
ovEU1 = product()
ovEU3 = product()
ovNL1 = product()
ovMOD = product()

#Independer
inEU1 = product()
inEU3 = product()
inNL1 = product()
inMOD = product()

"""
-------------------------------------
Gaslicht.com
-------------------------------------
"""

if gaslicht == "y":
    driver.get('https://www.gaslicht.com/energievergelijker')
    driver.maximize_window()
    time.sleep(3)
    #Searches for the cookiebox

    time.sleep(5)
    try:
        cookiesBox = driver.find_element_by_id("CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll")
        if cookiesBox.is_displayed(): #if displayed click the box
            cookiesBox.click()
    except:
        print("PASS - COOKIEBOX")
        pass

    #Searches for the cookiebox
    try:
        clickLabel = driver.find_element_by_id("CybotCookiebotDialogBodyLevelButtonAccept")
        if clickLabel.is_displayed(): #if displayed click the box
            clickLabel.click()
    except:
        print("PASS - COOKIEBOX 2 ")
        pass
    try:
        openOptions = driver.find_element_by_css_selector(".c-gl-compare-widget__usage-link.c-link.c-link--cta.js-compare-widget-trigger.custom")
        openOptions.click()
    except:
        print("PASS - COMPARE WIDGET")
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
        print("PASS - GEBRUIK ZELF INVULLEN")
        pass

    time.sleep(5)

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

    try:
        vergelijkPrijs = driver.find_element_by_xpath("/html/body/section/section[1]/div[2]/div/div/aside/div[1]/div/div/div[2]/form/div/div[2]/button")
        vergelijkPrijs.click()
    except:
        pass

    time.sleep(5)

    try:
      prijsdetails = driver.find_elements_by_xpath("//*[contains(text(), 'Prijsdetails')]")
      for x in range(0, len(prijsdetails)):
          if prijsdetails[x].is_displayed():
              prijsdetails[x].click()
              print("Displayed")
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

    glEU1.enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
    glEU1.teruglever = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
    glEU1.gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text

    #Het tweede element
    NED1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a')
    NED1Jaar = NED1JaarElement.text
    print(NED1Jaar)

    glNL1.enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
    glNL1.teruglever = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
    glNL1.gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text

    #Het derde element
    EU3JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a')
    EU3Jaar = EU3JaarElement.text
    print(EU3Jaar)

    glEU3.enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
    glEU3.teruglever = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
    glEU3.gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text


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

    time.sleep(5)

    #1st Element
    glEU1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[1]/div[1]/div[1]/div/div[4]/a').text
    print(glEU1JaarElement)#Tijdelijk

    glEU1.normaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
    glEU1.dal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
    glEU1.terugleverNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
    glEU1.terugleverDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text

    #2nd Element
    NED1JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[2]/div[1]/div[1]/div/div[4]/a').text
    print(NED1JaarElement)#Tijdelijk

    glNL1.normaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[1]/div').text
    glNL1.dal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[2]/div').text
    glNL1.terugleverNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[3]/div').text
    glNL1.terugleverDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab2"]/div[1]/div[2]/div[4]/div').text

    EU3JaarElement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol[2]/li[3]/div[1]/div[1]/div/div[4]/a').text
    print(EU3JaarElement)#Tijdelijk

    glEU3.normaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[1]/div').text
    glEU3.dal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[2]/div').text
    glEU3.terugleverNormaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[3]/div').text
    glEU3.terugleverDal = driver.find_element_by_xpath('//*[@id="js-async-content-tab3"]/div[1]/div[2]/div[4]/div').text


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
      prijsdetails = driver.find_element_by_xpath("//*[contains(text(), 'Prijsdetails')]")
      prijsdetails.click()

    except:
        print("ERROR - PRIJSDETAILS")
        pass

    time.sleep(3)

    glModelELement = driver.find_element_by_xpath('//*[@id="js-async-content"]/div[2]/ol/li/div[1]/div[1]/div/div[4]/a').text
    print(glModelELement)

    glMOD.normaal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text
    glMOD.dal = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[2]/div').text
    glMOD.gas = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[3]/div').text
    glMOD.teruglever = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[4]/div').text

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

    glMOD.enkel = driver.find_element_by_xpath('//*[@id="js-async-content-tab1"]/div[1]/div[2]/div[1]/div').text



"""
-------------------------------------
Overstappen.nl 
-------------------------------------
"""

if overstappen == "y":
    driver.get('https://www.overstappen.nl/energie/vergelijken/')
    driver.maximize_window()
    time.sleep(3)

    #Eerste scherm
    driver.find_element_by_name('postcode').send_keys(postcode)
    driver.find_element_by_name('housenumber').send_keys(huisNr)
    button = driver.find_element_by_xpath('//*[@id="esos-widget"]/div/form/div/div/div[3]/div/button')
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

    ovInnovaCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/div[1]/div/div/div')
    driver.execute_script("arguments[0].click();", ovInnovaCheckbox) # 1 jaar

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
        ovEU1.enkel = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[2]/span[2]').text
        ovEU1.gas =  driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[3]/span[2]').text
        ovEU1.vastrecht =  driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[4]/span[2]').text
        ovEU1.teruglever =  "-"
    except:
        print("ERROR - GEGEVENS VERKRIJGEN // OV1J")

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
    try:
        ovAanbieder = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/div/div[1]')
        if ovAanbieder.is_displayed():
            ovAanbieder.click()
    except:
        pass

    time.sleep(3)

    ovInnovaCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/div[2]/div/div[8]/div/input')
    driver.execute_script("arguments[0].click();", ovInnovaCheckbox)

    time.sleep(2)

    #Meer informatie
    driver.find_element_by_xpath("//*[contains(text(), 'Alle info')]").click()
    time.sleep(1)

    # Tarieven
    driver.find_element_by_xpath("//*[contains(text(), 'Tarieven & kosten')]").click()
    time.sleep(5)

#---------------------------------------------------------

    try:

        ovEU1.normaal = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[3]/span[2]').text
        ovEU1.dal = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div[1]/div/div[2]/div/div[2]/div/div/div[1]/div[2]/span[2]').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN // OV1JDBL")



#----------------------------------------------------------
    #3 jaar dubbele meter

    ovCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/div[3]/div/div/div')
    driver.execute_script("arguments[0].click();", ovCheckbox)  # 3 jaar checkbox

    time.sleep(3)

    # Meer informatie
    driver.find_element_by_xpath("//*[contains(text(), 'Alle info')]").click()
    time.sleep(1)

    # Tarieven
    driver.find_element_by_xpath("//*[contains(text(), 'Tarieven & kosten')]").click()

    try:
        ovEU3.normaal = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div/div/div[1]/div[3]/span[2]').text
        ovEU3.dal = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div/div/div[1]/div[2]/span[2]').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN // OV3JDBL")


#-------------------------------------------------------

    #wijzig naar 3 jaar enkel

    wijzigButton = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[1]/div[2]/button')
    wijzigButton.click()

    ovDubbeleCheckbox = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div[2]/div[2]/div[3]/form/div[5]/div/div/input')
    driver.execute_script("arguments[0].click();", ovDubbeleCheckbox)

    wijzigButton = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div[2]/div[2]/div[3]/form/div[10]/div/button')
    wijzigButton.click()

#---------------------------------

    #3 jaar enkel

    ovCheckbox = driver.find_element_by_xpath(
        '//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/div[3]/div/div/div')
    driver.execute_script("arguments[0].click();", ovCheckbox)  # 3 jaar checkbox

    time.sleep(3)

    ovInnovaCheckbox = driver.find_element_by_xpath(
        '//*[@id="esos-content"]/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/div[2]/div/div[8]/div/input')
    driver.execute_script("arguments[0].click();", ovInnovaCheckbox)

    time.sleep(3)

    # Meer informatie
    driver.find_element_by_xpath("//*[contains(text(), 'Alle info')]").click()
    time.sleep(1)

    # Tarieven
    driver.find_element_by_xpath("//*[contains(text(), 'Tarieven & kosten')]").click()

    try:
        ovEU3.enkel = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div/div/div[1]/div[2]/span[2]').text
        ovEU3.gas = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div/div/div[1]/div[3]/span[2]').text
        ovEU3.vastrecht = driver.find_element_by_xpath('//*[@id="esos-content"]/div/div/div/div/div/div[2]/div[4]/div/div/div[2]/div/div[2]/div/div/div[1]/div[4]/span[2]').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN // OV3J")

"""
-------------------------------------
Independer.nl
-------------------------------------
"""

if independer == 'y':
    driver.get('https://www.independer.nl/energie/intro.aspx')
    driver.maximize_window()
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

    #info
    time.sleep(2)
    parent = driver.find_element_by_id("product_1509")
    print("gevonden")
    meerinfo = parent.find_element_by_class_name("link-plus")
    print("gevonden info")
    meerinfo.click()
    print("gelukt")
    time.sleep(2)
    prijsdetails = parent.find_element_by_xpath("//*[contains(text(), 'Prijsdetails')]")
    prijsdetails.click()

    #gegevens ----

    #gegevens ----

    #wijzig gegevens
    driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-zoekresultaat/div/div/div/aside/ind-zoekresultaat-filters-container/ind-energie-zoekresultaat-filters/div/ind-energie-looptijd-filter/ind-zoekresultaat-filter-group/div/div/div[1]').click()

    time.sleep(3)
    inCheckbox = driver.find_element_by_xpath('//*[@id="radio-14"]')
    driver.execute_script("arguments[0].click();", inCheckbox)

    time.sleep(2)
    inAllContracts = driver.find_element_by_xpath('/html/body/ind-app/ind-shell/div/ng-component/ind-zoekresultaat/div/div/div/section/ind-zoekresultaat-view-container/ind-button-switches/ind-button-switch[3]/label')
    driver.execute_script("arguments[0].click();", inAllContracts)

    time.sleep(2)
    #3 jaar
    parent = driver.find_element_by_id("product_1666")
    print("gevonden")
    meerinfo = parent.find_element_by_class_name('link-plus')
    meerinfo.click()
    time.sleep(2)
    prijsdetails = parent.find_element_by_xpath("//*[contains(text(), 'Prijsdetails')]")
    prijsdetails.click()

    #wijzig gegevens
    inCheckbox = driver.find_element_by_xpath('//*[@id="radio-17"]')
    driver.execute_script("arguments[0].click();", inCheckbox)

    inAllContracts = driver.find_element_by_xpath('//*[@id="switch-5"]')
    driver.execute_script("arguments[0].click();", inAllContracts)


    #model
    parent = driver.find_element_by_id('product_1335')
    print("gevonden")
    meerinfo = parent.find_element_by_class_name("link-plus")
    time.sleep(2)
    meerinfo.click()
    time.sleep(2)
    prijsdetails = parent.find_element_by_xpath("//*[contains(text(), 'Prijsdetails')]")
    prijsdetails.click()

"""
-------------------------------------
Pricewise.nl
-------------------------------------
"""


if pricewise == "y":

    driver.get("https://www.pricewise.nl/energie-vergelijken/")
    driver.maximize_window()

    #invullen gegevens
    driver.find_element_by_xpath('//*[@id="pc_false"]').send_keys(postcode)
    driver.find_element_by_xpath('//*[@id="hn_false"]').send_keys(huisNr)


    usgaeArea = driver.find_element_by_xpath('//*[@id="usageArea"]')
    if usgaeArea.is_displayed():
        pwCheckbox = driver.find_element_by_xpath('//*[@id="metertype_false"]')
        driver.execute_script("arguments[0].click();", pwCheckbox)

        pwCheckbox = driver.find_element_by_xpath('//*[@id="spchecked_false"]')
        driver.execute_script("arguments[0].click();", pwCheckbox)

        verbruik = driver.find_element_by_xpath('//*[@id="elecPeak_false"]')
        verbruik.clear()
        verbruik.send_keys(verbruikNormaalTarief)
        verbruik = driver.find_element_by_xpath('//*[@id="elecOffPeak_false"]')
        verbruik.clear()
        verbruik.send_keys(verbruikDalTarief)
        verbruik = driver.find_element_by_xpath('//*[@id="gas_false"]')
        verbruik.clear()
        verbruik.send_keys(verbruikGas)

        driver.find_element_by_xpath('//*[@id="elecPeaksp_false"]').send_keys(terugNormaalTarief)
        driver.find_element_by_xpath('//*[@id="elecOffPeaksp_false"]').send_keys(terugDaltarief)

        select = Select(driver.find_element_by_name('suppliers'))
        select.select_by_visible_text('Onbekend / Anders')

        time.sleep(3)
        driver.find_element_by_xpath('//*[@id="en_0_btn_cta"]').click()
    else:
        driver.find_element_by_xpath('//*[@id="en_0_btn_manusg"]').click() #klikt op de "vul meter typen in"

        time.sleep(2)

        pwCheckbox = driver.find_element_by_xpath('//*[@id="metertype_false"]') #dubbele of slimme meter
        driver.execute_script("arguments[0].click();", pwCheckbox)

        pwCheckbox = driver.find_element_by_xpath('//*[@id="spchecked_false"]') #zonnepanelekn
        driver.execute_script("arguments[0].click();", pwCheckbox)

        verbruik = driver.find_element_by_xpath('//*[@id="elecPeak_false"]')
        verbruik.clear()
        verbruik.send_keys(verbruikNormaalTarief)
        verbruik = driver.find_element_by_xpath('//*[@id="elecOffPeak_false"]')
        verbruik.clear()
        verbruik.send_keys(verbruikDalTarief)
        verbruik = driver.find_element_by_xpath('//*[@id="gas_false"]')
        verbruik.clear()
        verbruik.send_keys(verbruikGas)

        driver.find_element_by_xpath('//*[@id="elecPeaksp_false"]').send_keys(terugNormaalTarief)
        driver.find_element_by_xpath('//*[@id="elecOffPeaksp_false"]').send_keys(terugDaltarief)

        select = Select(driver.find_element_by_name('suppliers'))
        select.select_by_visible_text('Onbekend / Anders')

        time.sleep(3)
        driver.find_element_by_xpath('//*[@id="en_0_btn_cta"]').click()

    time.sleep(5)

    driver.find_element_by_xpath('//*[@id="Leverancier"]/button').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="Leverancier"]/div/div[17]/div[1]').click()

    time.sleep(3)

    #1 jaar
    driver.find_element_by_xpath('//*[@id="en_1_btn_open_1"]/span[2]').click()

    time.sleep(3)
    try:
        tooltipster = driver.find_element_by_class_name('tooltipseter-content')
        if tooltipster.is_displayed():
            driver.find_element_by_class_name('tt-close').click()
    except:
        print("PASS - TOOLTIPSTER")
        pass

    try:
        pw1JaarNormaal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[3]/td[2]/span').text
        pw1JaarDal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[6]/td[2]/span').text
        pw1JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[10]/td[2]/span').text
        time.sleep(2)
        pw1JaarTerugDal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[11]/td[2]/span').text
        pw1JaarGas = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/span').text
        pw1JaarLeveringStroom = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[12]/td[2]/span').text
        pw1JaarLeveringGas = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/span').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN // PW1JDBL")

    #3 Jaar
    driver.find_element_by_xpath('//*[@id="en_1_btn_open_3"]/span[2]').click()

    time.sleep(3)

    try:
        pw3JaarNormaal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[3]/td[2]/span').text
        pw3JaarDal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[3]/td[2]/span').text
        pw3JaarTerugNormaal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[10]/td[2]/span').text
        pw3JaarTerugDal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[11]/td[2]/span').text
        pw3JaarGas = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/span').text
        pw3JaarLeveringStroom = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[12]/td[2]/span').text
        pw3JaarLeveringGas= driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T2_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/span').text
    except:
        pw3JaarNormaal = "-"
        pw3JaarDal = "-"
        pw3JaarTerugNormaal = "-"
        pw3JaarTerugDal ="-"
        pw3JaarGas = "-"
        pw3JaarLeveringStroom ="-"
        pw3JaarLeveringGas = "-"
        print("ERROR - GEGEVENS VERKRIJGEN // PW3JDBL")


    print(pw3JaarNormaal, pw3JaarDal, pw3JaarTerugNormaal, pw3JaarTerugDal, pw3JaarGas, pw3JaarLeveringStroom, pw3JaarLeveringGas)

    #Model
    driver.find_element_by_xpath('//*[@id="en_1_btn_open_4"]/span[2]').click()

    time.sleep(3)

    try:
        pwModelNormaal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[3]/td[2]/span').text
        pwModelDal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[6]/td[2]/span').text
        pwModelTerugNormaal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[10]/td[2]/span').text
        pwModelTerugDal = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[11]/td[2]/span').text
        pwModelGas = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/span').text
        pwModelLeveringStroom = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[12]/td[2]/span').text
        pwModelLeveringGas = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T2_C2-false"]/div[4]/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/span').text

        print(pwModelNormaal, pwModelDal, pwModelTerugNormaal, pwModelTerugDal, pwModelGas, pwModelLeveringStroom, pwModelLeveringGas)
    except:
        pwModelNormaal = "-"
        pwModelDal = "-"
        pwModelTerugNormaal = "-"
        pwModelTerugDal = "-"
        pwModelGas = "-"
        pwModelLeveringStroom = "-"
        pwModelLeveringGas = "-"
        print("ERROR - GEGEVENS VERKRIJGEN // PWMDLDBL")


    #wijzig

    driver.find_element_by_xpath('//*[@id="mainForm"]/div[3]/div/div/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div[7]/div/a').click() #wijzig button
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="metertype_true"]').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="mp_body"]/div[9]/div/div/div[3]/a[1]').click()

    #zonder dubele meter
    time.sleep(5)

    #1 jaar
    driver.find_element_by_xpath('//*[@id="en_1_btn_open_1"]/span[2]').click()
    try:
        pw1JaarEnkel = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T1_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[1]/td[2]/span').text
        pw1JaarTerugEnkel = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022596234_PG1022596235_T1_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[8]/td[2]/span').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN PW1JNKL")
        pw1JaarEnkel = "-"
        pw1JaarTerugEnkel = "-"

    # 3 Jaar
    driver.find_element_by_xpath('//*[@id="en_1_btn_open_3"]/span[2]').click()
    try:
        pw3JaarEnkel = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T1_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[1]/td[2]/span').text
        pw3JaarTerugEnkel = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1022605708_PG1022605709_T1_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[8]/td[2]/span').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN PW3JNKL")
        pw3JaarEnkel = "-"
        pw3JaarTerugEnkel = "-"


    # Model
    driver.find_element_by_xpath('//*[@id="en_1_btn_open_4"]/span[2]').click()
    try:
        pwModelEnkel = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T1_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[1]/td[2]/span').text
        pwModelTerugEnkel = driver.find_element_by_xpath('//*[@id="scrollto-M1_S1136_PE1006996_PG1006995_T1_C2-false"]/div[5]/div[3]/div/div/div[3]/div/div[1]/table/tbody/tr[8]/td[2]/span').text
    except:
        print("ERROR - GEGEVENS VERKRIJGEN PWMDLNKL")
        pwModelEnkel = "-"
        pwModelTerugEnkel = "-"

else:
    print('Klaar')


#To EXCEl

#EU1
tableEU1 = pd.DataFrame(data={
                              'Enkel':[glEU1.enkel, pwEU1.enkel, ovEU1.enkel, inEU1.enkel],
                              'Normaal':[glEU1.normaal, pwEU1.normaal, ovEU1.normaal, inEU1.normaal],
                              'Dal':[glEU1.dal, pwEU1.dal, ovEU1.dal, inEU1.dal],
                              'Gas':[glEU1.gas, pwEU1.gas, ovEU1.gas, inEU1.gas],
                              'Vastrecht Stroom':[glEU1.vastrecht, pwEU1.vastrecht, ovEU1.vastrecht, inEU1.vastrecht],
                              'Vastrecht gas':[glEU1.vastrecht, pwEU1.vastrecht, ovEU1.vastrecht, inEU1.vastrecht],
                              'Teruglevertarief':[glEU1.teruglever, pwEU1.teruglever, ovEU1.teruglever, inEU1.teruglever],
                              'Teruglevertarief Normaal':[glEU1.terugleverNormaal, pwEU1.terugleverNormaal, ovEU1.terugleverNormaal, inEU1.terugleverNormaal],
                              'Teruglevertarief Dal':[glEU1.terugleverDal, pwEU1.terugleverDal, ovEU1.terugleverDal, inEU1.terugleverDal]},
                             index=['Gaslicht.com', 'Pricewise.nl', 'Overstappen.nl', 'Independer.nl' ])
tableEU1.to_excel(writerTarieven, sheet_name='EU1', index=True,)
writerTarieven.save()

#EU3
tableEU3 = pd.DataFrame(data={
                              'Enkel':[glEU3.enkel, pwEU3.enkel, ovEU3.enkel, inEU3.enkel],
                              'Normaal':[glEU3.normaal, pwEU3.normaal, ovEU3.normaal, inEU3.normaal],
                              'Dal':[glEU3.dal, pwEU3.dal, ovEU3.dal, inEU3.dal],
                              'Gas':[glEU3.gas, pwEU3.gas, ovEU3.gas, inEU3.gas],
                              'Vastrecht Stroom':[glEU3.vastrecht, pwEU3.vastrecht, ovEU3.vastrecht, inEU3.vastrecht],
                              'Vastrecht gas':[glEU3.vastrecht, pwEU3.vastrecht, ovEU3.vastrecht, inEU3.vastrecht],
                              'Teruglevertarief':[glEU3.teruglever, pwEU3.teruglever, ovEU3.teruglever, inEU3.teruglever],
                              'Teruglevertarief Normaal':[glEU3.terugleverNormaal, pwEU3.terugleverNormaal, ovEU3.terugleverNormaal, inEU3.terugleverNormaal],
                              'Teruglevertarief Dal':[glEU3.terugleverDal, pwEU3.terugleverDal, ovEU3.terugleverDal, inEU3.terugleverDal]},
                             index=['Gaslicht.com', 'Pricewise.nl', 'Overstappen.nl', 'Independer.nl' ])
tableEU3.to_excel(writerTarieven, sheet_name='EU3', index=True,)
writerTarieven.save()

#NL1
tableNL1 = pd.DataFrame(data={
                              'Enkel':[glNL1.enkel, pwNL1.enkel, ovNL1.enkel, inNL1.enkel],
                              'Normaal':[glNL1.normaal, pwNL1.normaal, ovNL1.normaal, inNL1.normaal],
                              'Dal':[glNL1.dal, pwNL1.dal, ovNL1.dal, inNL1.dal],
                              'Gas':[glNL1.gas, pwNL1.gas, ovNL1.gas, inNL1.gas],
                              'Vastrecht Stroom':[glNL1.vastrecht, pwNL1.vastrecht, ovNL1.vastrecht, inNL1.vastrecht],
                              'Vastrecht gas':[glNL1.vastrecht, pwNL1.vastrecht, ovNL1.vastrecht, inNL1.vastrecht],
                              'Teruglevertarief':[glNL1.teruglever, pwNL1.teruglever, ovNL1.teruglever, inNL1.teruglever],
                              'Teruglevertarief Normaal':[glNL1.terugleverNormaal, pwNL1.terugleverNormaal, ovNL1.terugleverNormaal, inNL1.terugleverNormaal],
                              'Teruglevertarief Dal':[glNL1.terugleverDal, pwNL1.terugleverDal, ovNL1.terugleverDal, inNL1.terugleverDal]},
                             index=['Gaslicht.com', 'Pricewise.nl', 'Overstappen.nl', 'Independer.nl' ])
tableNL1.to_excel(writerTarieven, sheet_name='NL1', index=True,)
writerTarieven.save()

#MODEL
tableMOD = pd.DataFrame(data={
                              'Enkel':[glMOD.enkel, pwMOD.enkel, ovMOD.enkel, inMOD.enkel],
                              'Normaal':[glMOD.normaal, pwMOD.normaal, ovMOD.normaal, inMOD.normaal],
                              'Dal':[glMOD.dal, pwMOD.dal, ovMOD.dal, inMOD.dal],
                              'Gas':[glMOD.gas, pwMOD.gas, ovMOD.gas, inMOD.gas],
                              'Vastrecht Stroom':[glMOD.vastrecht, pwMOD.vastrecht, ovMOD.vastrecht, inMOD.vastrecht],
                              'Vastrecht gas':[glMOD.vastrecht, pwMOD.vastrecht, ovMOD.vastrecht, inMOD.vastrecht],
                              'Teruglevertarief':[glMOD.teruglever, pwMOD.teruglever, ovMOD.teruglever, inMOD.teruglever],
                              'Teruglevertarief Normaal':[glMOD.terugleverNormaal, pwMOD.terugleverNormaal, ovMOD.terugleverNormaal, inMOD.terugleverNormaal],
                              'Teruglevertarief Dal':[glMOD.terugleverDal, pwMOD.terugleverDal, ovMOD.terugleverDal, inMOD.terugleverDal]},
                             index=['Gaslicht.com', 'Pricewise.nl', 'Overstappen.nl', 'Independer.nl' ])
tableMOD.to_excel(writerTarieven, sheet_name='MODEL', index=True,)
writerTarieven.save()






