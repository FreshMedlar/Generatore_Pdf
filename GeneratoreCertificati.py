from codicefiscale import codicefiscale
from datetime import datetime
import os, sys
import tkinter as tk
from tkinter import messagebox
from docxtpl import DocxTemplate # docx
import xlwings as xw # excel 

import win32com.client as win32

template = "Attestato base logo SIF.docx"

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

def birth_format(codice_fiscale): # obtain the birthdate and format it to dd/mm/yyyy
    date = str(codicefiscale.decode(codice_fiscale)['birthdate']).partition(" ")[0].replace("-", "/").partition("/")
    year = date[0]
    month = date[2].partition('/')[0]
    day = date[2].partition('/')[2]
    return f'{day}/{month}/{year}'

def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", r".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None

def error_window(field): # show missing field 
    if len(field) == 1:
        messagebox.showerror('Error', f'Campo {field} mancante')
    else:
        messagebox.showerror('Error', 'Molteplici campi mancanti')



# excel selection and dictionary
def initial_function(excel):

    word_gen, pdf_gen = 1, 1 # control the generation of excel and word

    # open first sheet of excel and find the first valid row
    wb = xw.Book(excel)
    sht = wb.sheets[0]

    # finding the row with "nome" in it, where data start
    first_column = sht.range('A1').expand('table').value
    
    print(first_column)
    for row in range(len(first_column)):
        for cell in range(len(first_column[row])):
            if first_column[row][cell] == 'nome':
                initial_row = row
                initial_column = cell

    
    first_column_address = xw.Range(f'A{str(initial_row+2)}').expand("down").get_address(False, False)

    first_row = sht.range(f'A{str(initial_row+1)}').expand("right").value # chiavi del dizionario
    
    value_range = sht.range(f'A{str(initial_row+2)}').expand('table').value # all value

    # 0-nome
    # 1-cognome
    # 2-codice fiscale
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
    missing_field = []

    for i in range(len(value_range)):
        for j in range(len(first_row)):
            try:
                context[first_row[j].replace(" ", "")] = value_range[i][j]
            except IndexError:
                context[first_row[j].replace(" ", "")] = ""

        # report missing field
        missing_field.append(first_row[j])
        error_window(missing_field)

        context["data_nascita"] = birth_format(context["codicefiscale"])
        context["citta"] = codicefiscale.decode(context["codicefiscale"])["birthplace"]["name"]

        # replace pic
        doc.replace_pic('Picture 13', logo) # logo
        doc.replace_pic('Picture 4', firma) # firma

        # save docx
        output_name = f'{str(context["nome"]).replace(" ", "")}_{str(context["cognome"])}_{str(context["project"])}_{str(context["edition"])}.docx' # emanuele_segatori_project_edtition.docx
        doc.render(context)
        if word_gen:
            doc.save(output_name)

        # Convert to PDF
        path_word = os.path.join(os.getcwd(), output_name)
        if pdf_gen:
            convert_to_pdf(path_word)

if 0:
    initial_function("additional material/MODELLO FARC partecipanti_da_importare.xlsx")