# CODICE FISCALE
from codicefiscale import codicefiscale
from datetime import datetime
# OFFICE
from docxtpl import DocxTemplate # docx
import xlwings as xw # excel 
# GUI
import tkinter as tk
from tkinter.filedialog import askopenfilename
# PDF
import win32com.client as win32

import GeneratoreCertificati as gen
import os, sys

generate_file = True

# def save_setting(field):
#     f = open('setting.txt', 'w')
#     f.write('')

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
    if generate_file:
        gen.initial_function(excel)

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

root = tk.Tk()

# window creation
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