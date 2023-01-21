# -*- coding: utf-8 -*-
"""
Created on Sun Dec  8, 07:19 2020
SMK ROZKURWIATOR 0.6
Zmiany: 
    -dodana obsluga danych osob asystujacych
@author: Samuel Mazur
"""

import os
import sys
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlrd
import tkinter


def arkusz(): 
    lista = os.listdir('./arkusz')
    duzaDf = pd.DataFrame()
    for l in range(len(lista)):
        try:
            xls_file = pd.ExcelFile(os.path.join('./arkusz', lista[l]))
        except ValueError:
            print("Cannot read file "+lista[l], file=sys.stderr)
            continue
        df = xls_file.parse(0)
        duzaDf = duzaDf.append(df, ignore_index=True)
        duzaDf = duzaDf.astype(str)
    if 'Asysta' in duzaDf.columns:
        duzaDf = duzaDf[['Nazwisko', 'Imię', 'Usługa', 'Data opisu badania', 'Asysta']]
        duzaDf.insert(2, 'Plec', '0')
        duzaDf.insert(5, 'Inicjały', '0')
    else: 
        duzaDf = duzaDf[['Nazwisko', 'Imię', 'Usługa', 'Data opisu badania']]
        duzaDf.insert(2, 'Plec', '0')
        duzaDf.insert(5, 'Inicjały', '0')
        duzaDf.insert(6, 'Asysta', "")
    
    for i in range(duzaDf.shape[0]):
        #konwersja daty
        duzaDf.iat[i,4] = duzaDf.iat[i,4][0:10]
        #obciecie whitespace i smieci po imieniu
        head, sep, tail = duzaDf.iat[i,1].partition(' ')
        duzaDf.iat[i,1] = head
        #wytworzenie inicjalow
        duzaDf.iat[i,5] = duzaDf.iat[i,0][0] + '.' + duzaDf.iat[i,1][0] + '.'
        #identyfikacja plci
        if duzaDf.iat[i, 1].endswith("A"): 
            duzaDf.iat[i, 2] = 1
        #wywalenie nan z asysty
        if duzaDf.iat[i, 6]=='nan': duzaDf.iat[i,6] = ""
    return duzaDf
    
 
def dzialanie(table, rok, kod, nazwisko, miejsce, nazwa, xpath):
    for i in range(table.shape[0]):
        #driver.find_element_by_xpath('//button[@title="Dodaj"]').click()
        #wait.until(EC.element_to_be_clickable((By.XPATH, "//body[1]/div[3]/div[4]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/fieldset[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]/table[1]/tbody[1]/tr[5]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/button[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
    for i in range(table.shape[0]):
        #data
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[2]/div[1]/input[1]"))).send_keys(table.iat[i,4])
        #rok szkolenia
        rokSzkolenia = Select(wait.until(EC.element_to_be_clickable((By.XPATH,"//tbody/tr[" + str(i+1) + "]/td[4]/div[1]/select[1]"))))
        rokSzkolenia.select_by_value(rok)
        #kod zabiegu    
        kodZabiegu = Select(wait.until(EC.element_to_be_clickable((By.XPATH,"//tbody/tr[" + str(i+1) + "]/td[5]/div[1]/select[1]"))))
        kodZabiegu.select_by_index(str(int(kod)-1))
        #nazwisko
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[6]/div[1]/input[1]"))).send_keys(nazwisko)
        #miejsce
        miejscestazu_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i + 1) + "]/td[7]/div[1]/select[1]")))
        miejscestazu = Select(miejscestazu_element)
        try:
            miejscestazu.select_by_index(miejsce)
        except selenium.common.exceptions.WebDriverException:
            for j in range(int(miejsce)):
                miejscestazu_element.send_keys(Keys.ARROW_DOWN)
        #nazwastazu
        nazwaStazu_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[8]/div[1]/select[1]")))
        nazwaStazu = Select(nazwaStazu_element)
        try:
            nazwaStazu.select_by_index(nazwa)
        except selenium.common.exceptions.WebDriverException:
            for j in range(int(nazwa)):
                nazwaStazu_element.send_keys(Keys.ARROW_DOWN)
        #inicjaly
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[9]/div[1]/input[1]"))).send_keys(table.iat[i,5]) 
        #plec
        plec = Select(wait.until(EC.element_to_be_clickable((By.XPATH,"//tbody/tr[" + str(i+1) + "]/td[10]/div[1]/select[1]"))))
        if table.iat[i,2]==1: plec.select_by_value('K') 
        else: plec.select_by_value('M')
        #asysta
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[11]/div[1]/input[1]"))).send_keys(table.iat[i,6])
        #nazwaproc
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody/tr[" + str(i+1) + "]/td[12]/div[1]/input[1]"))).send_keys(table.iat[i,3]) 

class Okno(tkinter.Tk):
    def __init__(self, parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.grid()
        stepOne = tkinter.LabelFrame(self, text=" Wypełnij zgodnie z instrukcją ")
        stepOne.grid(row=0, columnspan=7, sticky='W',padx=5, pady=5, ipadx=5, ipady=5)
        self.RokLbl = tkinter.Label(stepOne,text="Rok szkolenia")
        self.RokLbl.grid(row=0, column=0, sticky='E', padx=5, pady=2)
        self.RokTxt = tkinter.Entry(stepOne)
        self.RokTxt.grid(row=0, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.KodLbl = tkinter.Label(stepOne,text="Kod zabiegu (1/2)")
        self.KodLbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)
        self.KodTxt = tkinter.Entry(stepOne)
        self.KodTxt.grid(row=1, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.OsobaLbl = tkinter.Label(stepOne,text="Osoba wykonująca")
        self.OsobaLbl.grid(row=2, column=0, sticky='E', padx=5, pady=2)
        self.OsobaTxt = tkinter.Entry(stepOne)
        self.OsobaTxt.grid(row=2, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.MiejsceLbl = tkinter.Label(stepOne,text="Miejsce wykonania (które miejsce na liscie)")
        self.MiejsceLbl.grid(row=3, column=0, sticky='E', padx=5, pady=2)
        self.MiejsceTxt = tkinter.Entry(stepOne)
        self.MiejsceTxt.grid(row=3, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.NazwaLbl = tkinter.Label(stepOne,text="Nazwa stażu (które miejsce na liscie)")
        self.NazwaLbl.grid(row=4, column=0, sticky='E', padx=5, pady=2)
        self.NazwaTxt = tkinter.Entry(stepOne)
        self.NazwaTxt.grid(row=4, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.XpathLbl = tkinter.Label(stepOne,text="Xpath")
        self.XpathLbl.grid(row=5, column=0, sticky='E', padx=5, pady=2)
        self.XpathTxt = tkinter.Entry(stepOne)
        self.XpathTxt.grid(row=5, column=1, columnspan=3, pady=2, sticky='WE')
        
        self.rok = None
        self.kod = None
        self.osoba = None
        self.miejsce = None
        self.nazwa = None
        self.xpath = None

        GuzikWysylania = tkinter.Button(stepOne, text="Wyslij",command=self.wyslij)
        GuzikWysylania.grid(row=6, column=3, sticky='W', padx=5, pady=2)

    def wyslij(self):
        self.rok = self.RokTxt.get()
        if self.rok == "":
            Win2=tkinter.Tk()
            Win2.withdraw()

        self.kod = self.KodTxt.get()
        if self.kod == "":
            Win2=tkinter.Tk()
            Win2.withdraw()
            
        self.osoba = self.OsobaTxt.get()
        if self.osoba == "":
            Win2=tkinter.Tk()
            Win2.withdraw()

        self.miejsce = self.MiejsceTxt.get()
        if self.miejsce == "":
            Win2=tkinter.Tk()
            Win2.withdraw()
        
        self.nazwa = self.NazwaTxt.get()
        if self.nazwa == "":
            Win2=tkinter.Tk()
            Win2.withdraw()
        
        self.xpath = self.XpathTxt.get()
        if self.xpath == "":
            Win2=tkinter.Tk()
            Win2.withdraw()
            
        self.quit()
        

#def main():
options = Options()
options.add_argument("start-maximized")
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")
options.add_argument("--log-level=3")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
try:
    driver = webdriver.Chrome(".\\chromedriver.exe")
except selenium.common.exceptions.WebDriverException:
    driver = webdriver.Firefox()
driver.maximize_window()
wait = WebDriverWait(driver, 50, poll_frequency=1)
driver.get("https://smk.ezdrowie.gov.pl/login.jsp?locale=pl")
tabela = arkusz()
app = Okno(None)
app.title("SMK Rozkurwiator 0.6")
app.mainloop() 
arg = [app.rok, app.kod, app.osoba, app.miejsce, app.nazwa, app.xpath]
dzialanie(tabela, *arg)
input()
# if __name__ == "__main__":
#     main()
