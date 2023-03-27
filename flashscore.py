from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver 
import os
options = webdriver.ChromeOptions()
driver=undetected_chromedriver.Chrome()
import pandas as pd
import glob


from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


name=" "
price=" "
priceSkidka=" "
nds=" "
KomercialType=" "
brendItem=" "
typeItem=" "
obiem=" "
articul= ''
width=''
#name,price,priceSkidka,nds,KomercialType
def parser(driver,id):
      
    global name, price,priceSkidka,nds ,KomercialType,brendItem,typeItem,obiem,countItem,femal,classop,ManufacturerСountry,articul,width    
    time.sleep(5)
    # переход  на сайт  ozon
    html=driver.get("https://www.ozon.ru/")
    time.sleep(5)
    #ищем  поле поиска
    kcal = driver.find_element(By.XPATH,"//input[contains(@class,'tsBodyL')]")
    time.sleep(5)
    #вставляем  id
    kcal.send_keys(id)
    time.sleep(5)
    #лупа для продолжения поиска
    lupa = driver.find_element(By.XPATH,'//button[@class="x5-a1 x5-f0"]')
    time.sleep(5)
    lupa.click()

    time.sleep(5)

    # клик по ссылке  продукта 
    item=driver.find_element(By.XPATH,"//span[contains(@class,'tsBodyL')]")
    item.click()
    time.sleep(5)
    

    
    
    # ищем имя  товара
    
    try: 
        name=driver.find_element(By.XPATH,"//h1[contains(@class,'rn')]").text
        name=name
        print(name)
    
    except:
        name=''
        print('элемент name(имя) не найден')     
    
    #ищем цену  товара
    try:
        price=driver.find_element(By.XPATH,"//span[contains(@class,'p1n np2')]").text.replace("₽", "")
        price=price
        print(price)
        
    except:
        price=''
        print('элемент price(цена) не найден')
    
    #ищем цену  до скидки товара
    try:
        priceSkidka=driver.find_element(By.XPATH,"//span[contains(@class,'pn2')]").text.replace("₽", "") 
        priceSkidka=priceSkidka 
        print(priceSkidka)      
    except:
        priceSkidka=''
        print('элемент priceSkidka(цена со скидкой) не найден')   
          
    #ищем  nds
    nds='13%'
    nds=nds
    print(nds)
 

    #KomercialType
    try:
        KomercialType=driver.find_element(By.XPATH,"//span[text()='Тип']/following::dd[1]").text
        KomercialType=KomercialType
        print(KomercialType)
    except:
        KomercialType=''
        print('элемент KomercialType не найден')
    
    #Ширина
    try:
        #r=re.match(\d[0-9])
        width=driver.find_element(By.XPATH,"//span[text()='Размер упаковки']/following::dd[1]").text
        width=width
        print(width)
    except:
        width=''
        print('элемент width не найден')
    #высота
    try:
        #r=re.match(\d[0-9])
        hidth=driver.find_element(By.XPATH,"//span[text()='Размер упаковки']/following::dd[1]").text
        hidth=hidth
        print(hidth)
    except:
        hidth=''
        print('элемент hidth не найден')
    #Длина
    try:
        #r=re.match(\d[0-9])
        width=driver.find_element(By.XPATH,"//span[text()='Размер упаковки']/following::dd[1]").text
        width=width
        print(width)
    except:
        width=''
        print('элемент width не найден')


   
    #SerialNumber
    #SerialNumber=driver.find_element(By.XPATH,"//span[text()='Тип']/following::dd[1]")
    #print(SerialNumber.text)
    #ищем  id 
    Ozonid=id
    print(Ozonid)
    #ищем тип  товара 
    #


    #ссылка на главное фото 
    #mainFoto=driver.find_element(By.XPATH,"//input[contains(@class,'tsBodyL')]")")
    #бренд   товара
    try:
        brendItem=driver.find_element(By.XPATH,"//span[text()='Бренд']/following::dd[1]").text
        brendItem= brendItem
        print(brendItem)
    except:
        brendIte=''
        print('элемент brendItem не найден')          
    time.sleep(45)

    #тип  товара
    #/td/span[text()='Данные2']/following::td[1]/span
    try:
        typeItem=driver.find_element(By.XPATH,"//span[text()='Тип']/following::dd[1]").text
        typeItem=typeItem
        print(typeItem)
    except:
        typeItem=''
        print('элемент typeItem не найден')

    #объем товара
    try:
        obiem=driver.find_element(By.XPATH,"//span[text()='Объем, мл']/following::dd[1]").text
        print(obiem)
    except:
        obiem=''
        print('элемент obiem не найден')
    #едениц товара
    countItem=1
    countItem=countItem
    print(countItem)  
    
    #объем товара
    try:
        articul=driver.find_element(By.XPATH,"//span[text()='Артикул']/following::dd[1]").text
        articul=articul
        print(articul)
    except:
        articul=''
        print('элемент articul не найден')
    
    #едениц товара
    countItem=1
    countItem=countItem
    print(countItem)  

    #пол
    try:
        femal=driver.find_element(By.XPATH,"//span[text()='Целевая аудитория']/following::dd[1]").text
        femal=femal
        print(femal)
    except:
        femal=''  
        print('Элемент femal не найден')
    #класс опасности
    classop="не опасен"
    classop=classop
    print(classop)
    

    #страна изготовитель
    try:
        ManufacturerСountry=driver.find_element(By.XPATH,"//span[text()='Страна-изготовитель']/following::dd[1]").text
        ManufacturerСountry=ManufacturerСountry
        print(ManufacturerСountry)
    except:
        ManufacturerСountry=''
        print('элемент ManufactureCountry не найден')    
       
    return id,name,price,priceSkidka,nds,KomercialType,articul
 
 


#,name,price,priceSkidka,nds,KomercialType
def OpenPixx(file,id,r,articul,name,price,priceSkidka,nds,KomercialType):
    hel='hello'
    wb = load_workbook(file)
    # get_sheet_names() - выводит список с названием листов,
    ws = wb['Шаблон для поставщика']
    #ws['B4']=id
    
    count=1
    ws.cell(row=r, column=1).value = count  #номер  столбца
    ws.cell(row=r, column=2).value = articul    #артикул
    ws.cell(row=r, column=3).value = name   #название товара
    ws.cell(row=r, column=4).value = price  #цена товара
    ws.cell(row=r, column=5).value = priceSkidka # цена до скидки
    ws.cell(row=r, column=6).value = nds
    ws.cell(row=r, column=7).value = id     #id
    ws.cell(row=r, column=8).value = KomercialType
    ws.cell(row=r, column=9).value = None
    ws.cell(row=r, column=10).value = obiem
    ws.cell(row=r, column=11).value = KomercialType #вес упаковки
    ws.cell(row=r, column=12).value = KomercialType #ширина упаковки
    ws.cell(row=r, column=12).value = KomercialType #высота упаковки
    wb.save(file)
    return 1

id=215973958    
parser(driver,id)
print(name)
print(ManufacturerСountry)

file ="C:\\ozon1\\Парфюмерия.xlsx"
#print(name,price,priceSkidka,nds,KomercialType)
#openExel(file)
r=9
c=3
OpenPixx(file,id,r,articul,name,price,priceSkidka,nds,KomercialType)

