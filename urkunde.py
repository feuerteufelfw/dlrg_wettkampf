import datetime
import os
import sqlite3
import time

import pythoncom

import main
import mailmerge
import win32com.client
from PyPDF2 import PdfMerger
from docx2pdf import convert
def mains():
    datenbank = sqlite3.connect("ergebnis.db")
    cursor = datenbank.cursor()
    disziplinen = main.get_disziplinen()
    ak = main.get_ak()
    #print(disziplinen)
    for akl in ak:
        for disisziplin in disziplinen:
            tabelle = akl + '_' +disisziplin
            print(tabelle)
            cursor.execute("SELECT * FROM " + tabelle)
            teilnehmer_list = cursor.fetchall()
            if len(teilnehmer_list) > 0:
                #print('zeile 209')
                #print(teilnehmer_list)
                write_to_docx( teilnehmer_list,disisziplin)
    print('schlafe 10 s')
    time.sleep(10)
    for f in os.listdir(os.path.abspath(".") + '/files/temp/docx'):
        test(os.path.abspath(".") + '/files/temp/docx/' + f)
        time.sleep(2)
    print('hiiiiiiiiiiiiiiiiiiiiiiiiii')
    #test('C:\\Users\\feuer\\\OneDrive\\\Desktop\\\dlrg_wettkampf\\files\\temp\\docx\\dokumen49.docs')
      #docx_to_pdf()
    merge_to_pdf()

def write_to_docx(list_teilnehmer,ak):
    print('write do dox')
    print(list_teilnehmer)
    check = True
    x = 1
    y = 1
    if len(list_teilnehmer) > 0:
        x = 0
        # erstellt dynamisch bis zu 10 dictionary die jeweils die daten zu einen Teilnehmer erhalten diese werden dann zusammen in eine word datei geschrieben
        # dies ist effizienter als jeden teilnehmer einzelt in eine word datei zu schreiben
        # eventuell erweiterung in 10ner schritte zur steigerung der effiziens
        for i in list_teilnehmer:
            x = x + 1
            #print(x)
            #print(i)
            datum = datetime.date.today()
            #teilnehmer_time = main.auswertung.decode_time(main.auswertung, i[5]).split(':')[2]
            #teilnehmer_time = str(teilnehmer_time)
            #teilnehmer_platzierung = i[6]
           # teilnehmer_platzierung = str(teilnehmer_platzierung)
            #print(i)
            globals()[f"cust_{x}"] = {
                'teilnehmer_Vorname': i[0],
                'teilnehmer_Nachname': i[1],
                'teilnehmer_ak':  ak,
#                'teilnehmer_Zeit': str( main.auswertung.decode_time(main.auswertung, i[5]).split(':')[2]),
                'datum': 'datums',
                'teilnehmer_platz': str(i[6])}
            if x == 10:
                urkunden_file = os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][
                    2] + '.docx'
                with mailmerge.MailMerge(urkunden_file) as Urkunden_dokument:
                    # print(urkunden_file)
                    Urkunden_dokument = mailmerge.MailMerge(urkunden_file)
                    print(cust_2)
                    print('Trage in docx  ein, 10 seiten')
                    Urkunden_dokument.merge_templates([cust_1, cust_2,cust_3,cust_4,cust_5,cust_6,cust_7,cust_8,cust_9,cust_10], separator='page_break')
                    print('daten übertargen')
                    x = 0
                    print('y: ' + str(y))
                    while os.path.isfile(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx'):
                        y = y + 1
                        print('y:' + str(y))
                    Urkunden_dokument.write(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx')

                    print('daten übertragen')
                    Urkunden_dokument.close()
                #test(urkunden_file)
        # print(x)
        urkunden_file = os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][2] + '.docx'
        # print(urkunden_file)
       # Urkunden_dokument = mailmerge.MailMerge(urkunden_file)
        with mailmerge.MailMerge(urkunden_file) as Urkunden_dokument:
            print(cust_1)
            if x == 10:
                # print('hi')
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
            # print(str(y))
            Urkunden_dokument.write(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx')
            print('daten übertragen')
            Urkunden_dokument.close()
        #test(urkunden_file)

    else:
        print('error keine daten erhalten')
    try:
        Urkunden_dokument.close()
    except:
        print('nö')
def docx_to_pdf():
    print('Convert docx to pdf')
    check = True
    x = 1

    for f in os.listdir(os.path.abspath(".") + '/files/temp/docx'):
        check = True
        while check:
            inputFile =os.path.abspath(".") + '/files/temp/docx/'+ f
            outputFile = os.path.abspath(".") + '/files/temp/pdf/' + str(x) + '.pdf'
            if os.path.isfile(outputFile):
                check = True
                x = x+1
            else:
                file = open(outputFile, "w")
                #file.close()
                print(inputFile)
                convert(inputFile, outputFile)
                file.close()
               # os.remove(inputFile)
                print(str(x))
                x = x+1
                check = False

def merge_to_pdf():
    print('merge pdf')
    check = True
    x = 1
    merger = PdfMerger()
    for f in os.listdir(os.path.abspath(".") + '/files/temp/pdf'):
        merger.append(os.path.abspath(".") + '/files/temp/pdf/'+f)
        print(f)
    merger.write(os.path.abspath(".") + '\\files\\test.pdf')
    merger.close()
def test(inputFile):
    print('start test')
    print('INput File: ' + inputFile)
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
            print(168)
    word.Visible = True
    print(170)
    print(inputFile)
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(str(outputFile), FileFormat=wdFormatPDF)
    doc.Close(0)
    word.Quit()

#if __name__ == '__main__':
 #   main()
#    print('done')
##test()
#docx_to_pdf()0
#merge_to_pdf()