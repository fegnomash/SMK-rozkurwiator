# -*- coding: utf-8 -*-
"""
Created on Mon Dec  7 11:53:23 2020
SMK Rozkurwiator dyżurów
@author: Samuel Mazur
"""

import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import xlrd
import tkinter

def arkusz(): 
    lista = os.listdir('.\\arkusz')
    duzaDf = pd.DataFrame()
    for l in range(len(lista)): 
        xls_file = pd.ExcelFile(os.path.join('.\\arkusz', lista[l]))
        df = xls_file.parse(0)
        duzaDf = duzaDf.append(df, ignore_index=True)
        duzaDf = duzaDf.astype(str)
    
    duzaDf = duzaDf[['Liczba godzin', 'Liczba minut', 'Data rozpoczęcia', 'Nazwa komórki organizacyjnej']]
    
    for i in range(duzaDf.shape[0]):
        #konwersja daty
        duzaDf.iat[i,3] = duzaDf.iat[i,3][0:10]
    return duzaDf

def dzialanie(table, xpath):
    for i in range(table.shape[0]):
        #driver.find_element_by_xpath('//button[@title="Dodaj"]').click()
        #wait.until(EC.element_to_be_clickable((By.XPATH, "//body[1]/div[3]/div[4]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/fieldset[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]/table[1]/tbody[1]/tr[5]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/button[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
    for i in range(table.shape[0]):
        #godziny
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[3]/div[1]/input[1]"))).send_keys(table.iat[i,0])
        #minuty
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[4]/div[1]/input[1]"))).send_keys(table.iat[i,1])
        #data
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[5]/div[1]/input[1]"))).send_keys(table.iat[i,2])
        #Nazwa komórki organizacyjnej
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[7]/div[1]/input[1]"))).send_keys(table.iat[i,3])

class Okno(tkinter.Tk):
    def __init__(self, parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.grid()
        stepOne = tkinter.LabelFrame(self, text=" Wypełnij zgodnie z instrukcją ")
        stepOne.grid(row=0, columnspan=7, sticky='W',padx=5, pady=5, ipadx=5, ipady=5)
        
        self.XpathLbl = tkinter.Label(stepOne,text="Xpath")
        self.XpathLbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)
        self.XpathTxt = tkinter.Entry(stepOne)
        self.XpathTxt.grid(row=1, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.xpath = None
        
        GuzikWysylania = tkinter.Button(stepOne, text="Wyslij",command=self.wyslij)
        GuzikWysylania.grid(row=2, column=3, sticky='W', padx=5, pady=2)
        
    def wyslij(self):               
        self.xpath = self.XpathTxt.get()
        if self.xpath == "":
            Win2=tkinter.Tk()
            Win2.withdraw()
            
        self.quit()

#def main():
driver = webdriver.Chrome(".\\chromedriver.exe")
options = Options()
options.add_argument("start-maximized")
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--log-level=3")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
wait = WebDriverWait(driver, 50, poll_frequency=1)
driver.get("https://smk.ezdrowie.gov.pl/login.jsp?locale=pl")
tabela = arkusz()
app = Okno(None)
app.title("SMK Rozkurwiator Dyżurów 0.1")
app.mainloop() 
arg = [app.xpath]
dzialanie(tabela, *arg)
input()    
# if __name__ == "__main__":
#     main()

        
        
