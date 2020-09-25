from selenium import webdriver
from CONFIG import *
import pandas as pd
import os
import time


path = 'C:\\Users\\kajro\\Documents\\Innova\\Pythonscripts\\Tarievenvergelijker\\'
chromedriverPath = path + 'chromedriver.exe'
driver = webdriver.Chrome(chromedriverPath)
driver.get('https://www.gaslicht.com/energievergelijker')

tarievenPath = path + 'tarieven.xlsx'

tableTarieven = pd.DataFrame(columns=['Vergelijker', 'Enkel', 'Normaal', 'Dal', 'Gas', 'Vastrecht Stroom', 'Vastrecht gas', 'Teruglevertarief'])
writerTarieven = pd.ExcelWriter(tarievenPath, engine='xlsxwriter')

tableTarieven = tableTarieven.append({'Vergelijker': 'Gaslicht.com'}, ignore_index=True)
tableTarieven = tableTarieven.append({'Vergelijker': 'Pricewise.nl'}, ignore_index=True)
tableTarieven = tableTarieven.append({'Vergelijker': 'Overstappen.nl'}, ignore_index=True)
tableTarieven = tableTarieven.append({'Vergelijker': 'Independer.nl'}, ignore_index=True)

tableEUR1Jaar = tableTarieven
tableEUR3Jaar = tableTarieven
tableNED1Jaar = tableTarieven
tableModel = tableTarieven







#naar excel:
#tableTarieven.to_excel(writerTarieven, sheet_name='EUR_1Jaar', index=False)
#writerTarieven.save()