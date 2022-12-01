import datetime
import os
import sqlite3
import time
import pythoncom
import main
import mailmerge
import win32com.client
from PyPDF2 import PdfMerger
def mains():
    datenbank = sqlite3.connect("ergebnis.db")
    cursor = datenbank.cursor()
    disziplinen = main.get_disziplinen()
    ak = main.get_ak()
    print(disziplinen)
    for akl in ak:
        for disisziplin in disziplinen:
            tabelle = akl + '_' +disisziplin
            print(tabelle)
            cursor.execute("SELECT * FROM " + tabelle)
            teilnehmer_list = cursor.fetchall()
            if len(teilnehmer_list) > 0:
                    #print('zeile 209')
                #print(teilnehmer_list)
                write_to_docx( teilnehmer_list,disisziplin,akl)
    print('schlafe 10 s')
    time.sleep(10)
    for f in os.listdir(os.path.abspath(".") + '/files/temp/docx'):
        docx_to_pdf(os.path.abspath(".") + '/files/temp/docx/' + f)
        time.sleep(1)
    merge_pdf()
def write_to_docx(list_teilnehmer,disziplin,ak):
    print('write do dox: ' + disziplin)
    print(list_teilnehmer)
    x = 1
    y = 1
    if len(list_teilnehmer) > 0:
        x = 0
        # erstellt dynamisch bis zu 10 dictionary die jeweils die daten zu einen Teilnehmer erhalten diese werden dann zusammen in eine word datei geschrieben
        # dies ist effizienter als jeden teilnehmer einzelt in eine word datei zu schreiben
        # eventuell erweiterung in 10ner schritte zur steigerung der effiziens
        datum = get_datum()
        for i in list_teilnehmer:
            x = x + 1
            teilnehmer_time = main.auswertung.decode_time(main.auswertung, i[5]).split(':')[2]
            teilnehmer_time = str(teilnehmer_time)
            globals()[f"cust_{x}"] = {
                'teilnehmer_Vorname': i[0],
                'teilnehmer_Nachname': i[1],
                'teilnehmer_ak':  ak,
                'teilnehmer_Zeit': teilnehmer_time,
                'datum': datum,
                'teilnehmer_platz': str(i[6])}
            if x == 10:
                urkunden_file = os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][
                    2] + '.docx'
                with mailmerge.MailMerge(urkunden_file) as Urkunden_dokument:
                    Urkunden_dokument.merge_templates([cust_1, cust_2,cust_3,cust_4,cust_5,cust_6,cust_7,cust_8,cust_9,cust_10], separator='page_break')
                    x = 0
                    while os.path.isfile(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx'):
                        y = y + 1
                    Urkunden_dokument.write(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx')
                    Urkunden_dokument.close()
        urkunden_file = os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][2] + '.docx'
        with mailmerge.MailMerge(urkunden_file) as Urkunden_dokument:
            print(cust_1)
            if x == 10:
                Urkunden_dokument.merge_templates(
                    [cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9, cust_10],
                    separator='page_break')
            elif x == 9:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9],
                                                  separator='page_break')
            elif x == 8:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8],
                                                  separator='page_break')
            elif x == 7:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7],
                                                  separator='page_break')
            elif x == 6:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6], separator='page_break')
            elif x == 5:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5], separator='page_break')
            elif x == 4:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4], separator='page_break')
            elif x == 3:
                Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3], separator='page_break')
            elif x == 2:
                Urkunden_dokument.merge_templates([cust_1, cust_2], separator='page_break')
            elif x == 1:
                Urkunden_dokument.merge_templates([cust_1], separator='page_break')
            Urkunden_dokument.write(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx')
            print('daten in docx übertragen')
            Urkunden_dokument.close()

    else:
        print('error keine daten erhalten')
    try:
        Urkunden_dokument.close()
    except:
        print('nö')
def get_datum():
    print('start get_datum')
    datum = datetime.date.today()
    tag = datum.day
    monat = datum.month
    jahr = datum.year
    datum = str(tag) + '.' + str(monat) + '.' + str(jahr)
    return datum
def merge_pdf():
    print('merge pdf')
    check = True
    x = 1
    merger = PdfMerger()
    for f in os.listdir(os.path.abspath(".") + '/files/temp/pdf'):
        merger.append(os.path.abspath(".") + '/files/temp/pdf/'+f)
        print(f)
    merger.write(os.path.abspath(".") + '\\files\\ergebnisse.pdf')
    merger.close()
def docx_to_pdf(inputFile):
    print('start docx_to_pdf')
    wdFormatPDF = 17
    x = 1
    word = win32com.client.DispatchEx("Word.Application",pythoncom.CoInitialize())
    word.Visible = True
    check = True
    while check:
        outputFile = os.path.abspath(".") + '/files/temp/pdf/' + str(x) + '.pdf'
        if os.path.isfile(outputFile):
            check = True
            x = x+1
        else:
            check = False
    word.Visible = True
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(str(outputFile), FileFormat=wdFormatPDF)
    doc.Close(0)
    word.Quit()