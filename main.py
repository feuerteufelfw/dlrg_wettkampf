#import Zeug
import os
import pythoncom
import win32com
from PyPDF2 import PdfMerger    #zusammenfügen von pdf dateien
import mailmerge #um auf die Textfelder in word zuzugreifen
import pandas as pd
    #word zu pdf
import Web_interface
import time #fürs zeitstoppen
import csv
import sqlite3
import datetime
#from subprocess import  Popen

import urkunde_old

#setup der files

#definiren einiger globaler variablen
wdFormatPDF = 17
t = 1
start_time = ''
#zum speichern von variablen welche von überall abgerufen werden können
class speicher:


    def __init__(self):
        self.config_file =''
        self.urkunde_file = ''
        self.new_zeiten_file = ''
        self.new_teilnehmer_file=''
        self.urkunde_output_file=''
        self.zwischenspeicher_file=''
        self.Zeiten_file =''
        self.teilnehmer_file_excl=''
        self.temp_pfd_pfad = ''
        self.temp_docx_pfad =''
        self.export_pfad = ''
        self.urkunden_file =''
        self.ak =[]
speicher_class = speicher()
def loade_config():
    speicher.temp_pdf_pfad = os.path.abspath(".") + '/files/temp/pdf/'
    speicher.temp_docx_pfad = os.path.abspath(".") + '/files/temp/docx/'
    speicher.export_pfad = os.path.abspath(".") + '\\files\\ergebnisse.pdf'
    speicher.urkunden_file =  os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/'
    #config_file = os.path.abspath(".") + 'files/config.txt'
    Teilnehmer_file_excl = os.path.abspath(".") + '/files/Teilnehmer.xlsx'
    speicher.teilnehmer_file_excl = Teilnehmer_file_excl
    Zeiten_file = os.path.abspath(".") + '/files/zeiten.csv'
    speicher.Zeiten_file = Zeiten_file
    zwischenspeicher_file = os.path.abspath(".") + '/files/temp/pdf/'
    speicher.zwischenspeicher_file = zwischenspeicher_file
    urkunden_Output_file = os.path.abspath(".") + '/files/Urkunden_Gesamt.pdf'
    speicher.urkunde_output_file = urkunden_Output_file
    new_teilnehmer_file = os.path.abspath(".") + '/files/new_tn.xlsx'
    speicher.new_teilnehmer_file = new_teilnehmer_file
    new_zeiten_file = os.path.abspath(".") + '/files/zeiten.csv'
    speicher.new_zeiten_file=new_zeiten_file
class auswertung():
    def auswertung(self):
        ak = loade_ak()
        print('Auswertung start')
        #auslesen der einzlenen Tielnehmer Daten
        with open(os.path.abspath(".") + '/files/zeiten.csv') as csvdatei:
            dataframe1 = pd.read_csv(os.path.join(os.path.abspath(".") + '/files/Teilnehmer.csv'), sep=',',index_col=False)
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
            auswertung.ak_und_disziplin_zuordnung(self,list_teilnehmer_dict,ak)
            auswertung.arry_sort(self,ak)
            #export.export(export,ak)
    def eintrag_vorhanden(self,teilnehmer,name,cursor): #checkt ob teilnehmer bereits in db vorhanden
        sql_command = '''SELECT * FROM ''' +name +''' WHERE Teilnehmernummer = '''+ teilnehmer['teilnehmer_Nummer']
        #print(sql_command)
        cursor.execute(sql_command)
        temp = cursor.fetchall()
        #print(111)
        #print(temp)
        if len(temp) < 1:
            return False
        else:
            return True

    def ak_und_disziplin_zuordnung(self,teilnehmer_list,ak): #speichert teilnehmer in db, tabellenname abhänigig von ak und disziplin
        print('start ak und disziplin zuordnung')
        datenbank = sqlite3.connect("ergebnis.db")
        disziplinen = get_disziplinen()
        altersklassen = ak
        cursor = datenbank.cursor()
        print( altersklassen)
      #  datenbank.set_trace_callback(print)
        for disziplin in disziplinen:
            for ak in altersklassen:
                name = ak + '_' + disziplin
                sql_command ='''SELECT count(name) FROM sqlite_master WHERE type='table' AND name="''' +name+ '''"'''
                cursor.execute(sql_command)
                if cursor.fetchone()[0] < 1:
                    sql_command = """CREATE TABLE """ + name + """(Teilnehmer_Vorname VARCHAR(50), Teilnehmer_Nachname VARCHAR(25), Disziplin VARCHAR(10), Verein VARCHAR(50), Teilnehmernummer VARCHAR(4), Zeit VARCHAR(32),Position VARCHAR(20))"""
                    #print(sql_command)
                    cursor.execute(sql_command)
        for teilnehmer in teilnehmer_list:
            #print(teilnehmer)
            teilnehmer_ak = teilnehmer['teilnehmer_Altersklasse']
            teilnehmer_disziplin = teilnehmer['teilnehmer_Disziplin']
            name = teilnehmer_ak + '_' + teilnehmer_disziplin
            if auswertung.eintrag_vorhanden(self,teilnehmer,name,cursor) == False:
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
                print('Neuen Teilnehmer in Datenbank eingefügt ')
            #print(name)
        #time.sleep(1)
        datenbank.commit()
        datenbank.close()

    def arry_sort(self,ak): #sortiert die datenbank um platzierungen zu ermitteln
        datenbank = sqlite3.connect("ergebnis.db")
        disziplinen = get_disziplinen()
        altersklassen = ak
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
                #print(teilnehmer_list)
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
                    #print('Zeile 190')
                    #print(i)
                    cursor.execute("""INSERT INTO """ + tabelle + """ VALUES (?,?,?,?,?,?,?)""", (
                        i[0], i[1], i[2], i[3], i[4], i[5], str(l)
                    ))
                    l = l + 1
                    datenbank.commit()
        datenbank.close()


    def decode_time(self,zeit): # gibt zeit im Format Stunden:Minuten:Sekunden zurück
        minuten, seconds = divmod(float(zeit), 60)
        #delta_time = time.monotonic()
        hours, minutes = divmod(minuten, 60)
        zeit = str(hours) + ':' + str(minutes) + ':' +str(round(seconds))
        return zeit

def new_teilnehmer_file():
    print('start new Teilnehmer')
    dataframe1 = pd.read_excel('files/Teilnehmer.xlsx', index_col=False) #liest die Daten in ein pandas dataframe ein
    dataframe2 = pd.read_csv(os.path.abspath(".") + '/files/teilnehmer.csv',index_col = False,)
    teilnehmer_vorhanden = False
    # vergleicht alle eingetragenen daten um dopplungen zu vermeiden
    #ist der Teilnehmer noch nicht in der alten liste wird er hinzugefügt
    for i, row in dataframe1.iterrows():
        tn_number = row.get('Teilnehmer Nummer')
        print(tn_number)
        print(get_teilnehmer_infos(tn_number))
        if get_teilnehmer_infos(tn_number).empty:
            teilnehmer_vorhanden = False
        else:
            teilnehmer_vorhanden = True
        if not teilnehmer_vorhanden:#wenn der Teilnehmer neu ist
            with open(os.path.abspath(".") + '/files/teilnehmer.csv','a') as csv_datei:
                writer = csv.writer(csv_datei)
                print(row)
                list = [tn_number,row.get('Vorname'),row.get('Nachname'),row.get('Verein'),row.get('Altersklasse'),row.get('Disziplin')]#print(dataframe2)
                print(list)
                writer.writerow(list)
                csv_datei.close()
    #dataframe2.to_excel(speicher.teilnehmer_file_excl,index=False)#überträgt alle teilnehmer in den haupt file
    #os.remove(speicher.new_teilnehmer_file)#löscht den file in dem die neuen Teilnehmer standen
def new_tn(tn_vorname,tn_nachname,tn_ak,tn_disziplin,verein):
    dataframe1 = pd.read_csv(  os.path.join(os.path.abspath(".") + '/files/Teilnehmer.csv'),sep = ',',index_col=False)
    tn_anzahl = len(dataframe1)
    tn_anzahl = tn_anzahl + 1
    tn_info = get_teilnehmer_infos(tn_anzahl)
    while tn_info.empty == False:
        tn_anzahl = tn_anzahl + 1
        tn_info = get_teilnehmer_infos(tn_anzahl)
    with open(os.path.abspath(".") + '/files/teilnehmer.csv', 'a') as csv_datei:
        writer = csv.writer(csv_datei)
        list = [tn_anzahl, tn_vorname,tn_nachname,verein,tn_ak,tn_disziplin]
        print(list)
        writer.writerow(list)
        csv_datei.close()
    return tn_anzahl
#alles runt um Zeitstoppen
class stoppuhr:
    def __init__(self,disziplin):
        self.diszilpin = disziplin
        self.start_time = time.monotonic()
    def new_time(self,teilnehmer_nummer,start_time): #fügt eine neue Zeit hinzu
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
        return delta_time
    def start_stoppuhr(self):
        start_time = time.monotonic()
def reset(): # löscht alles dateien
    try:
        os.remove(os.path.abspath(".") + '/files/zeiten.csv')
    except:
        print('kein zeitenfile vorhanden')
    try:
        os.remove(os.path.abspath(".") + '/files/Teilnehmer.xlsx')
    except:
        print('kein Teilnehmerfile vorhanden')
    try:
        os.remove(os.path.abspath(".")+"/ergebnis.db")
    except:
        print('keine ergebins Datenbank vorhanden')
    reset_export()
    reset_temp()
    if os.path.exists(os.path.abspath('.')+'/Teilnehmer.xlsx'):
        os.remove(os.path.abspath('.')+'Teilnehmer.xlsx')
    if os.path.exists(os.path.abspath('.')+'/zeiten.csv'):
        os.remove(os.path.abspath('.')+'/zeiten.csv')

def reset_Urkuden():
    print('start reset urkunden')
    x = 0
    for files in os.listdir(os.path.abspath(".") + "/files/Urkunden_Zusammenfassung/"):
        os.remove(os.path.abspath(".") + "/files/Urkunden_Zusammenfassung/" + files)
        x = x + 1
    print('Es wurden ' + str(x) + 'files gelöscht')
def reset_temp():
    print('start reset temp')
    x = 0
    for files in os.listdir(os.path.abspath(".") + "/files/temp/"):
        os.remove(os.path.abspath(".") + "/files/temp/" + files)
        x = x + 1
    print('Es wurden ' + str(x) + 'files gelöscht')
def get_disziplinen(): #return list mit allen disziplinen
    disziplinen_list = []
    for f in os.listdir(os.path.abspath(".") + '/files/Urkunden_Zusammenfassung'):
        print(f)
        disziplinen_list.append(f.split('.')[0])
    print(disziplinen_list)
    return disziplinen_list
def loade_ak():#list alle vorhandenen Altersklassen aus und returnt diese als list

    dataframe1 = pd.read_excel(os.path.join(os.path.abspath(".") + '/files/Teilnehmer.xlsx'), engine='openpyxl', index_col=False)
    df_altersklassen = dataframe1['Altersklasse']
    print(df_altersklassen)
    altersklassen = []
    for ak in df_altersklassen:
        vorhanden = False
        for temp in altersklassen:
            if temp == ak:
                vorhanden = True
        if vorhanden == False:
            altersklassen.append(ak)
    return altersklassen
class export:
    #in der export classe sind alle methoden die für den export der Daten in pdf und xlsx benötigt werden
    def __int__(self):
        self.temp_pfd_pfad =  os.path.abspath(".") + '/files/temp/pdf/'
        self.temp_docx_pfad = os.path.abspath(".") + '/files/temp/docx/'
        self.export_pfad= os.path.abspath(".") + '\\files\\ergebnisse.pdf'
    def export(self,disziplin ):
        ak = loade_ak()
        datenbank = sqlite3.connect("ergebnis.db")
        cursor = datenbank.cursor()
        disziplinen = get_disziplinen()
        print(disziplinen)
        for akl in ak:
            tabelle = akl + '_' + disziplin
            print(tabelle)
            cursor.execute("SELECT * FROM " + tabelle)
            teilnehmer_list = cursor.fetchall()
            if len(teilnehmer_list) > 0:
                # print('zeile 209')
                # print(teilnehmer_list)
                self.write_to_docx(self,teilnehmer_list, disziplin, akl)
        print('schlafe 5 s')
        time.sleep(5)
        for f in os.listdir(os.path.abspath(".") + '/files/temp/docx/'):
            self.docx_to_pdf(self,os.path.abspath(".") + '/files/temp/docx/' + f)
            time.sleep(1)
        self.merge_pdf(self,disziplin)
        time.sleep(2)
        self.delete_temp_files(self)
        self.export_xlsx(self,disziplin)
    def write_to_docx(self,list_teilnehmer, disziplin, ak):#überträgt daten in textfieldes des docx dokumentes
        print('write do dox: ' + disziplin)
        print(list_teilnehmer)
        x = 1
        y = 1
        if len(list_teilnehmer) > 0:
            x = 0
            # erstellt dynamisch bis zu 10 dictionary die jeweils die daten zu einen Teilnehmer erhalten diese werden dann zusammen in eine word datei geschrieben
            # dies ist effizienter als jeden teilnehmer einzelt in eine word datei zu schreiben
            # eventuell erweiterung in 10ner schritte zur steigerung der effiziens
            datum = self.get_datum(self)
            for i in list_teilnehmer:
                x = x + 1
                teilnehmer_time = auswertung.decode_time(auswertung, i[5]).split(':')[2]
                teilnehmer_time = str(teilnehmer_time)
                globals()[f"cust_{x}"] = {
                    'teilnehmer_Vorname': i[0],
                    'teilnehmer_Nachname': i[1],
                    'teilnehmer_ak': ak,
                    'teilnehmer_Zeit': teilnehmer_time,
                    'datum': datum,
                    'teilnehmer_platz': str(i[6])}
                if x == 10:
                    urkunden_file =  os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][
                        2] + '.docx'
                    with mailmerge.MailMerge(urkunden_file) as Urkunden_dokument:
                        Urkunden_dokument.merge_templates(
                            [cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9, cust_10],
                            separator='page_break')
                        x = 0
                        while os.path.isfile(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx'):
                            y = y + 1
                        Urkunden_dokument.write(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx')
                        Urkunden_dokument.close()
            urkunden_file =os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/' + list_teilnehmer[0][2] + '.docx'
            with mailmerge.MailMerge(urkunden_file) as Urkunden_dokument:
                print(cust_1)
                if x == 10:
                    Urkunden_dokument.merge_templates(
                        [cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9, cust_10],
                        separator='page_break')
                elif x == 9:
                    Urkunden_dokument.merge_templates(
                        [cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8, cust_9],
                        separator='page_break')
                elif x == 8:
                    Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7, cust_8],
                                                      separator='page_break')
                elif x == 7:
                    Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6, cust_7],
                                                      separator='page_break')
                elif x == 6:
                    Urkunden_dokument.merge_templates([cust_1, cust_2, cust_3, cust_4, cust_5, cust_6],
                                                      separator='page_break')
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
                while os.path.isfile(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx'):
                    y = y + 1
                Urkunden_dokument.write(os.path.abspath(".") + '/files/temp/docx/dokument' + str(y) + '.docx')
                print('daten in docx übertragen')
                Urkunden_dokument.close()

        else:
            print('error keine daten erhalten')
        try:
            Urkunden_dokument.close()
        except:
            print('nö')

    def docx_to_pdf(self,inputFile): # convertiert docx datei zu pdf
        print('start docx_to_pdf')
        wdFormatPDF = 17
        x = 1
        word = win32com.client.DispatchEx("Word.Application", pythoncom.CoInitialize())
        word.Visible = True
        check = True
        while check: #um keine vorhandenen datei zu überschreiben
            outputFile =os.path.abspath(".") + '/files/temp/pdf/'+ str(x) + '.pdf'
            if os.path.isfile(outputFile):
                check = True
                x = x + 1
            else:
                check = False
        word.Visible = True
        doc = word.Documents.Open(inputFile)
        doc.SaveAs(str(outputFile), FileFormat=wdFormatPDF)
        doc.Close(0)
        word.Quit()
    def merge_pdf(self,disziplin):#fügt alle dateien in /files/temp/pdf/ zu einer pdf zusammen
        print('merge pdf')
        check = True
        x = 1
        merger = PdfMerger()
        vorhanden = False
        for f in os.listdir(os.path.abspath(".") + '/files/temp/pdf/'):
            merger.append(os.path.abspath(".") + '/files/temp/pdf/' + f)
            vorhanden = True
            print(f)
        if vorhanden: # falls keine dateien exestieren wird auch keine merge datei erstellt
            merger.write(os.path.abspath(".") + '/files/export/ergebnis_' + disziplin + '.pdf')
        merger.close()

    def get_datum(self): #gibt das aktuelle datum zurück
        print('start get_datum')
        datum = datetime.date.today()
        tag = datum.day
        monat = datum.month
        jahr = datum.year
        datum = str(tag) + '.' + str(monat) + '.' + str(jahr)
        return datum
    def delete_temp_files(self): #löscht die files in files/temp
        print("start delete_temp_files")
        counter = 0
        for file in os.listdir(os.path.abspath(".") + '/files/temp/pdf/'):
            os.remove(os.path.abspath(".") + '/files/temp/pdf/' + file)
            counter = counter + 1
        for file in os.listdir(os.path.abspath(".")+'/files/temp/docx/'):
            os.remove(os.path.abspath(".")+ '/files/temp/docx/' + file)
            counter = counter +1
        print("es wurden " + str(counter) + ' Dateien gelöscht')
    def export_xlsx(self,disziplin): # überträgt daten aus sqllite db in xlsx datei
        db = sqlite3.connect("ergebnis.db")
        cursor = db.cursor()
        altersklassen = loade_ak()
        for ak in altersklassen:
            tabelle = ak + '_' + disziplin
            cursor.execute("SELECT * FROM " + tabelle)
            columns = [desc[0] for desc in cursor.description]
            data = cursor.fetchall()
            if len(data) > 0:
                df = pd.DataFrame(list(data), columns=columns)
                writer = pd.ExcelWriter( os.path.join(os.path.abspath(".") + '/files/export/export' + tabelle + '.xlsx'))
                df.to_excel(writer, sheet_name=tabelle)
                writer.save()

def get_teilnehmer_infos(teilnehmer_nummer):
    print('start get teilnehmer infos')
    dataframe1 = pd.read_csv(  os.path.join(os.path.abspath(".") + '/files/Teilnehmer.csv'),sep = ',',index_col=False) #ruft Teilnehmer.xlsx als dataframe auf
    print(dataframe1)
    teilnehmer = dataframe1.loc[dataframe1['Teilnehmer Nummer'] == int(teilnehmer_nummer)]#sucht den Teilnehmer mit der entsprechenden Teilnehmern nummer raus
    return teilnehmer
def reset_export():
    print('start reset export')
    x = 0
    for files in os.listdir(os.path.abspath(".") + "/files/export/"):
        os.remove(os.path.abspath(".") + "/files/export/" + files)
        x = x + 1
    print('Es wurden ' + str(x) + 'files gelöscht')
def get_export_files(): #returnt alle files in /files/export/
    files = []
    for file in os.listdir(os.path.abspath(".") + '/files/export/'):
        files.append(file)
    return files
def test():
    dataframe2 = pd.read_excel(os.path.abspath(".") + '/files/Teilnehmer.xlsx', index_col=False, )
    tn_anzahl = len(dataframe2)
    tn_anzahl = tn_anzahl + 1
    tn_info = get_teilnehmer_infos(tn_anzahl)
    if tn_info.empty:
        print('empty')
def main():
    if __name__ == '__main__':
        #loade_ak()
        loade_config()
        #export
        #export.export_xlsx(export,"2200m")
        new_teilnehmer_file()
        #test()
        #Web_interface.start_Web_interface()
if __name__ == '__main__':
    main()
