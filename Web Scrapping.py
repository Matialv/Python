#importamos los paquetes 
from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import numpy as np
import xlwings as xw
from selenium.common.exceptions import NoSuchElementException
import getpass
import os 

#Establecemos el directorio de trabajo
#os.chdir("c:/Users/" + getpass.getuser() + "/Desktop")

#Definimos los parametros de busqueda
objeto = 'Championes'

#Definimos los filtros 
genero = 'Hombre'
marca = ['Nike', 'Adidas']
#, 'New Balance', 'Reebok', 'Under Armour'
condicion = 'Nuevo'
Estilo = 'Deportivo'

#Definimos el ejecutable de chrome
driver = webdriver.Chrome(executable_path=r"C:\Users\matia\OneDrive\Desktop\Programas y ejecutables\chromedriver.exe")
#driver.close()
#le pasamos el URL que el driver tiene que ejecutar
driver.get('https://www.mercadolibre.com.uy/')
#driver.find_element_by_xpath('/html/body/div[3]/div/button').click()

#Maximizamos la ventana 
driver.maximize_window()

#Generamos listas y Df vacios para completarlos con la iteracion
Nombres = []
Monedas = []
Precios = []

df = pd.DataFrame()
#ingresamos los parametros de busqueda 
for i in range(len(marca)):
    busqueda = driver.find_element_by_xpath('/html/body/header/div/form/input')
    busqueda.clear()
    busqueda.send_keys(objeto + " " + genero + " " + Estilo + " " + condicion + " "+ marca[i])
    busqueda.send_keys(Keys.ENTER)
    time.sleep(3)
    
    #nuevoLink = driver.find_element_by_xpath('/html/body/main/div/div/section/div[1]/div/div/div/div[3]/a[2]').get_attribute('href')
    #driver.get(nuevoLink)
    #time.sleep(3)
    #vista = driver.element.get_attribute("href")
    #Datos a extraer

    #Busqueda 1
    nombre0 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[1]/div/div/a/div/div[3]/h2').text 
    Nombres.append(nombre0);
    moneda0 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda0);
    precio0 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio0)
    
        #Busqueda 2
    try: 
        nombre1 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[2]/div/div/a/div/div[2]/h2').text 
    except  NoSuchElementException: 
        nombre1 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[2]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre1);
    moneda1 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda1);
    precio1 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio1)
    
        #Busqueda 3
    try:
        nombre2 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[3]/div/div/a/div/div[2]/h2').text 
    except  NoSuchElementException: 
        nombre2 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[2]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre2);
    moneda2 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[3]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda2);
    precio2 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[1]/li[3]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio2)
    
        #Busqueda 4 
    try: 
        nombre3 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[1]/div/div/a/div/div[3]/h2').text 
    except  NoSuchElementException:
        nombre3 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[1]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre3);
    moneda3 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda3);
    precio3 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio3)
    
        #Busqueda 5
    try:
        nombre4 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[2]/div/div/a/div/div[3]/h2').text 
    except  NoSuchElementException:
        nombre4 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[2]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre4);
    moneda4 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda4);
    precio4 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio4)
    
        #Busqueda 6
    try:
        nombre5 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[3]/div/div/a/div/div[2]/h2').text 
    except NoSuchElementException:
        nombre5 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[3]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre5);
    moneda5 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[3]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda5);
    precio5 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[2]/li[3]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio5)
    
        #Busqueda 7
    try: 
        nombre6 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[1]/div/div/a/div/div[2]/h2').text 
    except NoSuchElementException:
        nombre6 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[1]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre6);
    moneda6 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda6);
    precio6 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio6)
    
        #Busqueda 8
    try: 
        nombre7 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[2]/div/div/a/div/div[2]/h2').text 
    except NoSuchElementException:
        nombre7 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Nombres.append(nombre7);
    moneda7 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda7);
    precio7 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio7)
    
        #Busqueda 9
    try: 
        nombre8 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[3]/div/div/a/div/div[3]/h2').text 
    except NoSuchElementException:
        nombre8 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[3]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre8);
    moneda8 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[3]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda8);
    precio8 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[3]/li[3]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio8)
    
        #Busqueda 10
    try:  
        nombre9 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[4]/li[1]/div/div/a/div/div[2]/h2').text 
    except NoSuchElementException:
        nombre9 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[4]/li[1]/div/div/a/div/div[3]/h2').text
    Nombres.append(nombre9);
    moneda9 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[5]/li[2]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[1]').text
    Monedas.append(moneda9);
    precio9 = driver.find_element_by_xpath('/html/body/main/div/div/section/ol[4]/li[1]/div/div/a/div/div[1]/div/div/span[1]/span[2]/span[2]').text
    Precios.append(precio9)

      
    #Creamos un dataframe con los datos obtenidos 
    datos = {'Producto': Nombres, 'Moneda': Monedas, 'Precio' : Precios}
    df1 = pd.DataFrame(datos)
    df1.sort_values(by = ['Moneda', 'Precio'], ascending=True)
    df1.set_index('Producto')
    driver.back()
    
    
driver.close()   

#df1['Precio'].str.replace(",", ".").astype(int)

wb = xw.Book()
sht = wb.sheets['Hoja1']
sht.range('A1').value = df1
    
    
    
    
    
    
    
    
    

    

