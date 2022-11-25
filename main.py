#import Zeug
import os
from PyPDF2 import PdfMerger    #zusammenfügen von pdf dateien
import mailmerge #um auf die Textfelder in word zuzugreifen
import pandas as pd
from docx2pdf import convert    #word zu pdf
import Web_interface
import time #fürs zeitstoppen
import csv
import sqlite3
import datetime
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
    #if config_file :
    #    with open(config_file) as txt_config_datei:
    #        config_lines = txt_config_datei.readlines()
    #        for item in config_lines:
    #            print(item)
    #muss noch aus config datei geladen werden
    #übertrag in speicher class für leichteren zugriff
    stoppuhr.ort = 'Blossin'
    stoppuhr.veranstalter_vorname = 'Martin'
    stoppuhr.veranstalter_nachname = 'Krüger'
    #config_file = os.path.abspath(".") + 'files/config.txt'
    speicher.config_file = config_file
    urkunde_file = os.path.abspath(".") + 'files/Urkunde.docx'
    speicher.urkunde_file = urkunde_file
    Teilnehmer_file_excl = os.path.abspath(".") + '/files/Teilnehmer.xlsx'
    speicher.teilnehmer_file_excl = Teilnehmer_file_excl
    Zeiten_file = os.path.abspath(".") + '/files/zeiten.csv'
    speicher.Zeiten_file = Zeiten_file
    zwischenspeicher_file = os.path.abspath(".") + '/files/temp/pdf/'
    speicher.zwischenspeicher_file = zwischenspeicher_file
    urkunden_Output_file = os.path.abspath(".") + '/files/Urkunden_Gesamt.pdf'
    speicher.urkunde_output_file = urkunden_Output_file
    new_teilnehmer_file = os.path.abspath(".") + '/files/Teilnehmer.xlsx'
    speicher.new_teilnehmer_file = new_teilnehmer_file
    new_zeiten_file = os.path.abspath(".") + '/files/zeiten.csv'
    speicher.new_zeiten_file=new_zeiten_file
class auswertung():
    def auswertung(self,disziplin):
        print('Auswertung start')
        #auslesen der einzlenen Tielnehmer Daten
        with open(os.path.abspath(".") + '/files/zeiten.csv') as csvdatei:
            dataframe1 = pd.read_excel(os.path.abspath(".") + '/files/Teilnehmer.xlsx',index_col = False)
            csv_reader_object =csv.reader(csvdatei)
            y = 1
            x = 1
           # print(csv_reader_object)
            list_teilnehmer_dict = []
            for row in csv_reader_object:
               # print(row[0])
                teilnehmer_nummer = row[0]

                teilnehmer = dataframe1.loc[dataframe1['Teilnehmer Nummer'] == int(teilnehmer_nummer)]
               # print(teilnehmer)
                teilnehmer_Zeit = row[1]
                teilnehmer_Vorname = teilnehmer['Vorname'].values[0]
                teilnehmer_Nachname =  teilnehmer['Nachname'].values[0]
                teilnehmer_Verein = teilnehmer['Verein'].values[0]
                teilnehmer_Altersklasse = teilnehmer['Altersklasse'].values[0]
                teilnehmer_Disziplin = teilnehmer['Disziplin'].values[0]
                print(teilnehmer_Vorname)
               # print('vorname ' + teilnehmer['Vorname'])
                cust_1 = {
                    'teilnehmer_Vorname': teilnehmer_Vorname,
                    'teilnehmer_Nachname': teilnehmer_Nachname,
                    'teilnehmer_Altersklasse': teilnehmer_Altersklasse,
                    'teilnehmer_Verein': teilnehmer_Verein,
                    'teilnehmer_Disziplin': teilnehmer_Disziplin,
                    'teilnehmer_Zeit': teilnehmer_Zeit,
                    'teilnehmer_Nummer': teilnehmer_nummer,
                }
                list_teilnehmer_dict.append(cust_1)
                #if y == 10:
                #    auswertung.write_to_docx(list_teilnehmer_dict,x)
                #    list_teilnehmer_dict = []
                 #   y = 0
                  #  x = x + 1
                #y = y+1
            #if list_teilnehmer_dict:
            #    auswertung.write_to_docx(list_teilnehmer_dict, x)
            #auswertung.docx_to_pdf()
            #auswertung.merge_to_pdf(disziplin)
               # time.sleep(0.2)
            auswertung.ak_und_disziplin_zuordnung(self,list_teilnehmer_dict)
            auswertung.arry_sort(self)
            auswertung.export(self)

    def ak_und_disziplin_zuordnung(self,teilnehmer_list):
        datenbank = sqlite3.connect("ergebnis.db")
        disziplinen = get_disziplinen()
        altersklassen = get_ak()
        cursor = datenbank.cursor()
      #  datenbank.set_trace_callback(print)
        for disziplin in disziplinen:
            for ak in altersklassen:
                name = ak + '_' + disziplin
                sql_command ='''SELECT count(name) FROM sqlite_master WHERE type='table' AND name="''' +name+ '''"'''
                cursor.execute(sql_command)
                if cursor.fetchone()[0] < 1:
                    sql_command = """CREATE TABLE """ + name + """(Teilnehmer_Vorname VARCHAR(50), Teilnehmer_Nachname VARCHAR(25), Disziplin VARCHAR(10), Verein VARCHAR(50), Teilnehmernummer VARCHAR(4), Zeit FLOATE(32),Position VARCHAR(20))"""
                    print(sql_command)
                    cursor.execute(sql_command)
        #globals()[f"{disziplinen + '_' + altersklassen}"]
        for teilnehmer in teilnehmer_list:
            print(teilnehmer)
            teilnehmer_ak = teilnehmer['teilnehmer_Altersklasse']
            teilnehmer_disziplin = teilnehmer['teilnehmer_Disziplin']
            name = teilnehmer_ak + '_' + teilnehmer_disziplin
            sql_command="""INSERT INTO """ + name + """ VALUES (?,?,?,?,?,?,?)""",(
                teilnehmer['teilnehmer_Vorname'],
                teilnehmer['teilnehmer_Nachname'],
                teilnehmer['teilnehmer_Disziplin'],
                teilnehmer['teilnehmer_Verein'],
                teilnehmer['teilnehmer_Nummer'],
                teilnehmer['teilnehmer_Zeit'],
                'hi'

            )
            cursor.execute("""INSERT INTO """ + name + """ VALUES (?,?,?,?,?,?,?)""",(
                teilnehmer['teilnehmer_Vorname'],
                teilnehmer['teilnehmer_Nachname'],
                teilnehmer['teilnehmer_Disziplin'],
                teilnehmer['teilnehmer_Verein'],
                teilnehmer['teilnehmer_Nummer'],
                teilnehmer['teilnehmer_Zeit'],
                'hi'

            ))
            datenbank.commit()
            print(name)
        #time.sleep(1)
        datenbank.commit()
        datenbank.close()

    def arry_sort(self):
        datenbank = sqlite3.connect("ergebnis.db")
        disziplinen = get_disziplinen()
        altersklassen = get_ak()
        # datenbank.set_trace_callback(print)
        cursor = datenbank.cursor()
        for disziplin in disziplinen:
            for ak in altersklassen:
                tabelle = ak + '_' + disziplin
                cursor.execute("SELECT * FROM " + tabelle)
                teilnehmer_list = cursor.fetchall()
                #  print(teilnehmer_list)
                cursor.execute("DELETE FROM " + tabelle)
                datenbank.commit()

                teilnehmer_list.sort(key=lambda x: x[5])
                temp = []
                print(teilnehmer_list)
                for i in range(len(teilnehmer_list)):
                    for x in range(len(teilnehmer_list)):
                        if x > i:
                            if teilnehmer_list[i][5] > teilnehmer_list[x][5]:
                                temp = teilnehmer_list[i]
                                teilnehmer_list[i] = teilnehmer_list[x]
                                teilnehmer_list[x] = temp
                            elif teilnehmer_list[i][5] == teilnehmer_list[x][5]:
                                print('error gleiche zeit ')
                l = 1
                for i in teilnehmer_list:
                    print(i)
                    cursor.execute("""INSERT INTO """ + tabelle + """ VALUES (?,?,?,?,?,?,?)""", (
                        i[0], i[1], i[2], i[3], i[4], i[5], str(l)
                    ))
                    l = l + 1
                    datenbank.commit()
        datenbank.close()

    def export(self):
        disziplinen = get_disziplinen()
        alterstklassen = get_ak()
        datenbank = sqlite3.connect("ergebnis.db")
        cursor = datenbank.cursor()
        for disziplin in disziplinen:
            for ak in alterstklassen:
                tabelle = ak + '_' + disziplin
                cursor.execute("SELECT * FROM " + tabelle)
                teilnehmer_list = cursor.fetchall()
                if len(teilnehmer_list) > 0:
                    auswertung.write_to_docx(self,teilnehmer_list,tabelle,ak)
                    files = []
                    for f in os.listdir(os.path.abspath(".") + '/files/temp/'):
                        files.append(f)
                    auswertung.docx_to_pdf(self,files)
                    auswertung.merge_to_pdf()
                else:
                    print('keine daten erhalten')
    def merge_to_pdf(self):
        print('merge pdf')
        check = True
        x = 1
        merger = PdfMerger()
        for f in os.listdir(os.path.abspath(".") + '/files/temp/pdf'):
                merger.append(f)

        merger.write(os.path.abspath(".") + '/files')
        merger.close()
    def docx_to_pdf(self,files):
        print('Convert docx to pdf')
        check = True
        x = 1

        for f in files:
            while check:
                inputFile = f
                outputFile = os.path.abspath(".") + '/files/temp/pdf/' + str(x) + '.pdf'
                if os.path.isfile(outputFile):
                    check = True
                    x = x+1
                else:
                    file = open(outputFile, "w")
                    file.close()
                    convert(inputFile, outputFile)
                    print(str(x))
                    x = x+1
                    check = False
    def decode_time(self,zeit):
        minuten, seconds = divmod(zeit, 60)
        #delta_time = time.monotonic()
        hours, minutes = divmod(minuten, 60)
        zeit = str(hours) + ':' + str(minutes) + ':' +str(round(seconds))
        return zeit
    def write_to_docx(self,list_teilnehmer,y,ak,):
        print('write do dox')
        print(list_teilnehmer)
        check = True
        x = 1
        if len(list_teilnehmer) > 0:

            urkunden_file = os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][2] + '.docx'


            print(urkunden_file)
            Urkunden_dokument = mailmerge.MailMerge(urkunden_file)
            x = 0
            # erstellt dynamisch bis zu 10 dictionary die jeweils die daten zu einen Teilnehmer erhalten diese werden dann zusammen in eine word datei geschrieben
            # dies ist effizienter als jeden teilnehmer einzelt in eine word datei zu schreiben
            #eventuell erweiterung in 10ner schritte zur steigerung der effiziens
            for i in list_teilnehmer:
                x =x+1
                print(x)
                globals()[f"cust_{x}"] = {
                    'teilnehmer_Vorname': i[0],
                    'teilnehmer_Nachname': i[1],
                    'teilnehmer_Altersklasse': ak,
                    'teilnehmer_Verein': i[3],
                    'teilnehmer_Disziplin': i[2],
                    'teilnehmer_Zeit': auswertung.decode_time(self,i[5]).split(':')[1],
                    'datum': datetime.date.today()
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
            Urkunden_dokument.write(speicher.zwischenspeicher_file + y + '.docx')
        else:
            print('error keine daten erhalten')
def new_teilnehmer():
    print('start new Teilnehmer')
    dataframe1 = pd.read_excel(speicher.new_teilnehmer_file, index_col=False) #liest die Daten in ein pandas dataframe ein
    dataframe2 = pd.read_excel(os.path.abspath(".") + '/files/Teilnehmer.xlsx',index_col = False,)
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
        #minuten, seconds = divmod(delta_time, 60)
        #delta_time = time.monotonic()
        #hours, minutes = divmod(minuten, 60)
        #zeit = str(hours) + ',' + str(minutes) + ',' +str(seconds)
        ergebnis = teilnehmer_nummer,delta_time
        print(ergebnis)
        #speichert teilnehmernummer und die dazugehörige zeit in csv
        with open( os.path.abspath(".") + '/files/zeiten.csv', 'a',newline='') as csvfile_old_time:
            zeit_save = csv.writer(csvfile_old_time, quoting=csv.QUOTE_ALL)
            zeit_save.writerow(ergebnis)
    def start_stoppuhr(self):
        start_time = time.monotonic()
def reset():
    print('deleting all user data')
    try:
        os.remove(os.path.abspath(".") + '/files/zeiten.csv')
    except:
        print('kein zeitenfile vorhanden')
    try:
        os.remove(os.path.abspath(".") + '/files/Teilnehmer.xlsx')
    except:
        print('kein Teilnehmerfile vorhanden')

    for f in os.listdir(os.path.abspath(".") + '/files/Urkunden_Zusammenfassung'):
        os.remove(os.path.join(os.path.abspath(".") + '/files/Urkunden_Zusammenfassung', f))
def get_disziplinen():
    disziplinen_list = []
    for f in os.listdir(os.path.abspath(".") + '/files/Urkunden_Zusammenfassung'):
        print(f)
        disziplinen_list.append(f.split('.')[0])
    print(disziplinen_list)
    return disziplinen_list
def get_ak():
    return ["AK1","Ak2","Ak3","Ak4","Ak5"]
def main():
    if __name__ == '__main__':
        speicher
        loade_config()
        Web_interface.start_Web_interface()
if __name__ == '__main__':
    main()
