
from codicefiscale import codicefiscale

import os, sys

from docxtpl import DocxTemplate # docx
import xlwings as xw # excel 

import tkinter as tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32

# os.chdir(sys.path[0])
# excel = r'C:\Users\medlar\Desktop\progetto_gabellone\MODELLO FARC partecipanti_da_importare.xlsx'
# template = r'C:\Users\medlar\Desktop\progetto_gabellone\Attestato base logo SIF.docx'

# base bianca
logo = 'bianco.png'
firma =  'bianco.png' 
project = ""
release_date= ""
edition = ""
durata= ""
inizio = ""
fine  = ""
corso= ""


root = tk.Tk()

def choose_excel():
    global excel
    excel = askopenfilename(parent= root)

choose_excel()

# excel selection and dictionary
wb = xw.Book(excel)
sht = wb.sheets['elenco']

first_column = sht.range('A1').expand('down').value
print(first_column)
for i in range(len(sht.range('A1').expand('down').value)):
    if first_column[i] == 'nome':
        initial_row= i

print(initial_row)

first_row = sht.range(f'A{str(initial_row+1)}:C{str(initial_row+1)}').value # sht.range('A1').expand('right').value # chiavi del dizionario
value_range = sht.range(f'A{str(initial_row+2)}').expand('table').value # lista di rows dalla seconda

print(first_row)
print(value_range)
## 0-nome
## 1-cognome
## 2-codice fiscale
# 3-codice cittadinanza
# 4-codice livello studio
# 5-codice tipologia contrattuale
# 6-codice ccnl
# 7-codice inquadramento
# 8-anno anzianità
# 9-Assunzione ai sensi ex lege 68/99
# 10-cellulare
# 11-matricola INPS
# 12-cap (aggiornato)
# 13-città di residenza
# 14-indirizzo residenza
# 15-email
# 16-costo orario lordo azienda
