from codicefiscale import codicefiscale
from datetime import datetime
import os, sys
import tkinter as tk
from tkinter import messagebox
from docxtpl import DocxTemplate # docx
import xlwings as xw # excel 

import win32com.client as win32


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
def initial_function(excel, project, release_date, edition, durata, inizio, fine, corso, template, firma, logo):#, template_attestati, template_materiali):

    word_gen, pdf_gen = 1, 1 # control the generation of excel and word

    template_attestati = "Distinta consegna attestati Logo SIF.docx"
    template_materiali = "Distinta consegna materiali Logo SIF.docx"
    # open first sheet of excel and find the first valid row
    wb = xw.Book(excel)
    sht = wb.sheets[0]

    # finding the row with "nome" in it, where data start
    first_column = sht.range('A1').expand('table').value
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
    # 8-anno anzianit??
    # 9-Assunzione ai sensi ex lege 68/99
    # 10-cellulare
    # 11-matricola INPS
    # 12-cap (aggiornato)
    # 13-citt?? di residenza
    # 14-indirizzo residenza
    # 15-email
    # 16-costo orario lordo azienda

    doc_materiali = DocxTemplate(template_materiali)
    doc_attestati = DocxTemplate(template_attestati)
    doc = DocxTemplate(template)
    context = {"project":project, "release_date":release_date, "edition":edition, "corso":corso, "durata":durata, "inizio":inizio, "fine":fine}
    context_attestati = {}
    missing_field = []

    # folder creation
    path = os.getcwd()+f"\Attestati P{context['project']} Ed{context['edition']}"
    if not(os.path.exists(path)):
        os.mkdir(path)

    i = 0
    for i in range(len(value_range)):
        for j in range(len(first_row)):
            try:
                context[first_row[j].replace(" ", "")] = value_range[i][j]
            except IndexError:
                context[first_row[j].replace(" ", "")] = ""

        context_attestati["nome"+str(i)] = context["nome"]
        context_attestati["cognome"+str(i)] = context["cognome"]

        # report missing field
        # missing_field.append(first_row[j])
        # error_window(missing_field)

        context["data_nascita"] = birth_format(context["codicefiscale"])
        context["citta"] = codicefiscale.decode(context["codicefiscale"])["birthplace"]["name"]

        # replace pic
        doc.replace_pic('Picture 13', logo) # logo
        doc.replace_pic('Picture 4', firma) # firma
        
        # save docx
        output_name = f'{path}\{str(context["nome"]).replace(" ", "")}_{str(context["cognome"])}_{str(context["project"])}_{str(context["edition"])}.docx' # emanuele_segatori_project_edtition.docx
        doc.render(context)
        
        if word_gen:
            doc.save(output_name)
            doc_attestati.save("prova_attestati.docx")
            doc_materiali.save("prova_materiali.docx")

        # Convert to PDF
        path_word = os.path.join(os.getcwd(), output_name)
        
        if pdf_gen:
            convert_to_pdf(path_word)

    print(context_attestati)
    doc_attestati.render(context_attestati)
    doc_materiali.render(context_attestati)
    doc_attestati.save("prova_attestati.docx")
    doc_materiali.save("prova_materiali.docx")
    

if 0:
    initial_function("additional material/MODELLO FARC partecipanti_da_importare.xlsx")