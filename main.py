#import Zeug
import os
from PyPDF2 import PdfMerger    #zusammenfügen von pdf dateien
import mailmerge #um auf die Textfelder in word zuzugreifen
import pandas as pd
from docx2pdf import convert    #word zu pdf
import Web_interface
import time #fürs zeitstoppen
import csv
#setup der files

#definiren einiger globaler variablen
wdFormatPDF = 17
t = 1
start_time = ''
#zum speichern von variablen welche von überall abgerufen werden können
class speicher:
    def __init__(self):
        self.ort = ''
        self.veranstalter_vorname = ''
        self.veranstalter_nachname = ''
        self.config_file =''
        self.urkunde_file = ''
        self.new_zeiten_file = ''
        self.new_teilnehmer_file=''
        self.urkunde_output_file=''
        self.zwischenspeicher_file=''
        self.Zeiten_file =''
        self.teilnehmer_file_excl=''
        self.disziplinen_list = ["2000m","500m","1000m"]
def loade_config():
    config_file = os.path.abspath(".") + '\\config.txt'
    if config_file:
        with open(config_file) as txt_config_datei:
            config_lines = txt_config_datei.readlines()
            for item in config_lines:
                print(item)
    #muss noch aus config datei geladen werden
    #übertrag in speicher class für leichteren zugriff
    stoppuhr.ort = 'Blossin'
    stoppuhr.veranstalter_vorname = 'Martin'
    stoppuhr.veranstalter_nachname = 'Krüger'
    config_file = os.path.abspath(".") + 'files/config.txt'
    speicher.config_file = config_file
    urkunde_file = os.path.abspath(".") + 'files/Urkunde.docx'
    speicher.urkunde_file = urkunde_file
    Teilnehmer_file_excl = os.path.abspath(".") + '/files/Teilnehmer.xlsx'
    speicher.teilnehmer_file_excl = Teilnehmer_file_excl
    Zeiten_file = os.path.abspath(".") + '/files/zeiten.csv'
    speicher.Zeiten_file = Zeiten_file
    zwischenspeicher_file = os.path.abspath(".") + '/files/Urkunden_Zusammenfassung'
    speicher.zwischenspeicher_file = zwischenspeicher_file
    urkunden_Output_file = os.path.abspath(".") + '/files/Urkunden_Gesamt.pdf'
    speicher.urkunde_output_file = urkunden_Output_file
    new_teilnehmer_file = os.path.abspath(".") + '/files/Teilnehmer.xlsx'
    speicher.new_teilnehmer_file = new_teilnehmer_file
    new_zeiten_file = os.path.abspath(".") + '/files/zeiten.csv'
    speicher.new_zeiten_file=new_zeiten_file
def auswertung(disziplin):
    print('Auswertung start')
    #auslesen der einzlenen Tielnehmer Daten
    with open(speicher.Zeiten_file) as csvdatei:
        dataframe1 = pd.read_excel('Teilnehmer.xlsx',index_col = False)
        csv_reader_object =csv.reader(csvdatei)
        y = 1
        x = 1
       # print(csv_reader_object)
        list_teilnehmer_dict = []
        for row in csv_reader_object:
            row = row[0]
            teilnehmer_nummer = row.split('"')[0]
            teilnehmer_nummer = teilnehmer_nummer.split(',')[0]
            teilnehmer = dataframe1.loc[dataframe1['Teilnehmer Nummer'] == teilnehmer_nummer]
            teilnehmer_Zeit = row.split('"')[1]
            teilnehmer_Zeit = teilnehmer_Zeit.split(',')[0] + ':' + teilnehmer_Zeit.split(',')[1] + ':' + teilnehmer_Zeit.split(',')[2]
            teilnehmer_Vorname = teilnehmer['Vorname'].to_string(index = False)
            teilnehmer_Nachname =  teilnehmer['Nachname'].to_string(index = False)
            teilnehmer_Verein = teilnehmer['Verein'].to_string(index=False)
            teilnehmer_Altersklasse = teilnehmer['Altersklasse'].to_string(index=False)
            teilnehmer_Disziplin = teilnehmer['Disziplin'].to_string(index=False)
           # print('vorname ' + teilnehmer['Vorname'])
            cust_1 = {
                'teilnehmer_Vorname': teilnehmer_Vorname,
                'teilnehmer_Nachname': teilnehmer_Nachname,
                'teilnehmer_Altersklasse': teilnehmer_Altersklasse,
                'teilnehmer_Verein': teilnehmer_Verein,
                'teilnehmer_Disziplin': teilnehmer_Disziplin,
                'teilnehmer_Zeit': teilnehmer_Zeit,
                'ort': speicher.ort,
                'veranstalter_Vorname': speicher.veranstalter_vorname,
                'veranstalter_Nachname': speicher.veranstalter_nachname,
            }
            list_teilnehmer_dict.append(cust_1)
            if y == 10:
                write_to_docx(list_teilnehmer_dict,x)
                list_teilnehmer_dict = []
                y = 0
                x = x + 1
            y = y+1
        if list_teilnehmer_dict:
            write_to_docx(list_teilnehmer_dict, x)
        docx_to_pdf()
        merge_to_pdf(disziplin)
           # time.sleep(0.2)
def merge_to_pdf(disziplin):
    print('merge pdf')
    check = True
    x = 1
    merger = PdfMerger()
    while  check == True:
        if os.path.exists(os.path.abspath(".") + '\\temp\\document_' + str(x) + '.pdf'):
            print(str(x))
            merger.append(os.path.abspath(".") + '\\temp\\document_' + str(x) + '.pdf')
            x = x+1
        else:
            check = False
    merger.write(speicher.urkunden_Output_file)
    merger.close()
def docx_to_pdf():
    print('Convert docx to pdf')
    check = True
    x = 1
    while check == True:
        if os.path.exists(os.path.abspath(".") + '\\temp\\Urkunden_Zusammenfassung' + str(x) + '.docx'):
            print(str(x))
            inputFile = os.path.abspath(".") + '\\temp\\Urkunden_Zusammenfassung' + str(x) + '.docx'
            outputFile = os.path.abspath(".") + '\\temp\\document_' + str(x) + '.pdf'
            file = open(outputFile, "w")
            file.close()
            convert(inputFile, outputFile)
            print(str(x))
            x = x+1
        else:
            check = False
def write_to_docx(list_teilnehmer_dict,y):
    #print(list_teilnehmer_dict)
    Urkunden_dokument = mailmerge(speicher.urkunde_file)
    x = 0
    # erstellt dynamisch bis zu 10 dictionary die jeweils die daten zu einen Teilnehmer erhalten diese werden dann zusammen in eine word datei geschrieben
    # dies ist effizienter als jeden teilnehmer einzelt in eine word datei zu schreiben
    #eventuell erweiterung in 10ner schritte zur steigerung der effiziens
    for i in list_teilnehmer_dict:
        x =x+1
        globals()[f"cust_{x}"] = {
            'teilnehmer_Vorname': i['teilnehmer_Vorname'],
            'teilnehmer_Nachname': i['teilnehmer_Nachname'],
            'teilnehmer_Altersklasse': i['teilnehmer_Altersklasse'],
            'teilnehmer_Verein': i['teilnehmer_Verein'],
            'teilnehmer_Disziplin': i['teilnehmer_Disziplin'],
            'teilnehmer_Zeit': i['teilnehmer_Zeit'],
            'ort': i['ort'],
            'veranstalter_Vorname': i['veranstalter_Vorname'],
            'veranstalter_Nachname': i['veranstalter_Nachname']

        }
    #print(x)
    #print(cust_1)
    if x == 10 :
        #print('hi')
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9, cust_10],separator='page_break')
    elif x == 9 :
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9],separator='page_break')
    elif x == 8:
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8], separator='page_break')
    elif x == 7:
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7],separator='page_break')
    elif x == 6:
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6],separator='page_break')
    elif x == 5:
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5],separator='page_break')
    elif x == 4:
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4 ],separator='page_break')
    elif x == 3:
        Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3],separator='page_break')
    elif x == 2:
        Urkunden_dokument.merge_templates([cust_1, cust_2],separator='page_break')
    elif x == 1:
        Urkunden_dokument.merge_templates([cust_1],separator='page_break')
   # print(str(y))
    Urkunden_dokument.write(speicher.zwischenspeicher_file + str(y) + '.docx')
def new_teilnehmer():
    print('start new Teilnehmer')
    dataframe1 = pd.read_excel(speicher.new_teilnehmer_file, index_col=False) #liest die Daten in ein pandas dataframe ein
    dataframe2 = pd.read_excel(speicher.teilnehmer_file_excl,index_col = False,)
    teilnehmer_vorhanden = False
    # vergleicht alle eingetragenen daten um dopplungen zu vermeiden
    #ist der Teilnehmer noch nicht in der alten liste wird er hinzugefügt
    for i, row in dataframe1.iterrows():
        for x,row2 in dataframe2.iterrows():
            # print(row)
            #print('row1: ' + str(row.get('Teilnehmer Nummer')))
           # print('row2: ' + str(row2.get('Teilnehmer Nummer')))
            if row.get('Teilnehmer Nummer') == row2.get('Teilnehmer Nummer'):#wenn der Teilnehmer bereits vorhanden ist
                print('Teilnehmer ist bereits vorhanden')
                teilnehmer_vorhanden = True
        if not teilnehmer_vorhanden:#wenn der Teilnehmer neu ist
            #print('neuen Teilnehmer speichern')
            #print(row)
            #print(i)
            dataframe2 = dataframe2.append(row)
        teilnehmer_vorhanden = False
    #print(dataframe2)
    dataframe2.to_excel(speicher.teilnehmer_file_excl,index=False)#überträgt alle teilnehmer in den haupt file
    os.remove(speicher.new_teilnehmer_file)#löscht den file in dem die neuen Teilnehmer standen

def new_time():
    with open(speicher.new_zeiten_file,newline='') as csvfile_new_time:
        zeite_new = csv.reader(csvfile_new_time,delimiter = ',')
        neue_zeiten=[]
        with open (speicher.Zeiten_file,newline='') as csvfile_old_time:
            zeiten_old =csv.reader(csvfile_old_time)
            zeit_vorhanden = False
            for row_old_time in zeiten_old:
                for row_new_time in zeite_new:
            # print(row)
            # print('row1: ' + str(row.get('Teilnehmer Nummer')))
            # print('row2: ' + str(row2.get('Teilnehmer Nummer')))
                    if row_new_time == row_old_time:
                        print('Zeit ist bereits vorhanden')
                        zeit_vorhanden = True
                if not zeit_vorhanden:
                    print('neuen Teilnehmer speichern')
                    neue_zeiten.append(row_new_time)
            with open(speicher.Zeiten_file,'a', newline='') as csvfile_old_time:
                zeit_save = csv.writer(row_new_time)
                zeit_save.writerows(neue_zeiten)
            zeit_vorhanden = False;
        #print(dataframe2)
       # dataframe2.to_excel(Teilnehmer_file_excl, index=False)
        os.remove(speicher.new_zeiten_file)
#alles runt um Zeitstoppen
class stoppuhr:
    def __init__(self,disziplin):
        self.diszilpin = disziplin
        self.start_time = time.monotonic()
    def new_time(self,teilnehmer_nummer,start_time):
        print(teilnehmer_nummer)
       # print('start time: ' + strstart_time)
        delta_time = time.monotonic() - start_time
        #zeit wird in Stunden Minuten Sekunden umgerechnet
        minuten, seconds = divmod(delta_time, 60)
        delta_time = time.monotonic()
        hours, minutes = divmod(minuten, 60)
        zeit = str(hours) + ',' + str(minutes) + ',' +str(seconds)
        ergebnis = teilnehmer_nummer,zeit
        print(ergebnis)
        #speichert teilnehmernummer und die dazugehörige zeit in csv
        with open(speicher.Zeiten_file, 'a',newline='') as csvfile_old_time:
            zeit_save = csv.writer(csvfile_old_time, quoting=csv.QUOTE_ALL)
            zeit_save.writerow(ergebnis)

    def start_stoppuhr(self):
        start_time = time.monotonic()

def main():
    if __name__ == '__main__':
        Web_interface.start_Web_interface()
if __name__ == '__main__':
    main()
