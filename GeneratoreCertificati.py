from importlib.resources import path
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

# functions
def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", r".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None

def choose_excel():
    global excel
    excel = askopenfilename(parent= root, title="Modello base excel")
    
def choose_word():
    global template
    template = askopenfilename(parent = root, title="Modello base word")   

def choose_logo():
    global logo
    logo = askopenfilename(parent = root, title = "Logo eventuale")

def choose_firma():
    global firma
    firma = askopenfilename(parent = root, title = "Firma eventuale")

def confirm_project():
    global project 
    global release_date
    global edition 
    global durata
    global inizio 
    global fine 
    global corso
    project = entry_project.get()
    release_date = entry_release_date.get()
    edition = entry_edition.get()
    durata = entry_durata.get()
    inizio = entry_inizio.get()
    fine = entry_fine.get()
    corso = entry_corso.get()

# GUI
root = tk.Tk()

canvas = tk.Canvas(root, width=500, height = 400)
canvas.grid(columnspan = 8, rowspan = 9)

# button excel
excel_str = tk.StringVar()
excel_btn = tk.Button(root, textvariable= excel_str, command=lambda:choose_excel())
excel_str.set("Excel")
excel_btn.grid(column=0, row=0)

# button word
word_str = tk.StringVar()
word_btn = tk.Button(root, textvariable= word_str, command=lambda:choose_word())
word_str.set("Word")
word_btn.grid(column=0, row=1)

# button logo
logo_str = tk.StringVar()
logo_btn = tk.Button(root, textvariable= logo_str, command=lambda:choose_logo())
logo_str.set("logo")
logo_btn.grid(column=0, row=2)

# button firma
firma_str = tk.StringVar()
firma_btn = tk.Button(root, textvariable= firma_str, command=lambda:choose_firma())
firma_str.set("firma")
firma_btn.grid(column=0, row=3)

# entry codice progetto
entry_project = tk.Entry(root)
entry_project.grid(column=2, row=0)
project_lbl = tk.Label(root, text="Progetto")
project_lbl.grid(column=1, row=0)

# entry edizione
entry_edition = tk.Entry(root)
entry_edition.grid(column=2, row=1)
edition_lbl = tk.Label(root, text="Edizione")
edition_lbl.grid(column=1, row=1)

# entry data rilascio
entry_release_date = tk.Entry(root)
entry_release_date.grid(column=2, row=2)
release_lbl = tk.Label(root, text="Data di rilascio")
release_lbl.grid(column=1, row=2)

# entry corso
entry_corso = tk.Entry(root)
entry_corso.grid(column=2, row=3)
corso_lbl = tk.Label(root, text="Nome Corso")
corso_lbl.grid(column=1, row=3)

# entry durata
entry_durata = tk.Entry(root)
entry_durata.grid(column=2, row=4)
durata_lbl = tk.Label(root, text="Durata")
durata_lbl.grid(column=1, row=4)

# entry inizio
entry_inizio = tk.Entry(root)
entry_inizio.grid(column=2, row=5)
inizio_lbl = tk.Label(root, text="Data inizio corso")
inizio_lbl.grid(column=1, row=5)

# entry fine
entry_fine = tk.Entry(root)
entry_fine.grid(column=2, row=6)
fine_lbl = tk.Label(root, text="Data fine corso")
fine_lbl.grid(column=1, row=6)

# button confirm
entry_btn = tk.Button(root, text= "conferma", command= confirm_project)
entry_btn.grid(column=0, row=7)

root.mainloop()

# excel selection and dictionary
wb = xw.Book(excel)
sht = wb.sheets['elenco']

first_column = sht.range('A1').expand('down').value
print(first_column)
for i in range(len(sht.range('A1').expand('down').value)):
    if first_column[i] == 'nome':
        initial_row= i

breakpoint()

first_row = sht.range('A1:C1').value # sht.range('A1').expand('right').value # chiavi del dizionario
value_range = sht.range('A2').expand('table').value # lista di rows dalla seconda

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

doc = DocxTemplate(template)
context = {"project":project, "release_date":release_date, "edition":edition, "corso":corso, "durata":durata, "inizio":inizio, "fine":fine}


for i in range(len(value_range)):
    for j in range(len(first_row)):
        context[first_row[j].replace(" ", "")] = value_range[i][j]
    context["data_nascita"] = str(codicefiscale.decode(context["codicefiscale"])['birthdate']).replace("-", "/").partition(" ")[0]
    context["citta"] = codicefiscale.decode(context["codicefiscale"])["birthplace"]["name"]

    # replace pic
    print(logo, firma)
    doc.replace_pic('Picture 13', logo) # logo
    doc.replace_pic('Picture 4', firma) # firma

    # save docx
    output_name = f'{str(context["nome"]).replace(" ", "")}_{str(context["cognome"])}_{str(context["project"])}_{str(context["edition"])}.docx' 
    doc.render(context)
    doc.save(output_name)

    # Convert to PDF
    path_word = os.path.join(os.getcwd(), output_name)
    convert_to_pdf(path_word)
    


# name and save 
# output_name = 'prova.docx' # f'{str(context["nome"])}_{str(context["cognome"])}_progetto_edizione_.docx'  # f'{str(context[list(context.keys())[0]])}_{str(context[list(context.keys())[1]])}_progetto_edizione_.docx'
# doc.render(context)
# doc.save(output_name)