#import gedöns
import os
#import threading
from pathlib import Path
import flask
import main
from flask_dropzone import Dropzone
import Web_interface
import time
from flask_bootstrap import Bootstrap
import test
#____________________________________________________definition aller verwendeten pfade
basedir = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'files')
app = flask.Flask(__name__,static_folder='static')
Bootstrap(app)
time_file = os.path.join(basedir, 'time')
config_file = os.path.join(basedir, 'config')
teilnehmer_file = basedir + '/files'
urkunden_file = basedir  + "/files/Urkunden_Zusammenfassung"
visibility_urkunden = "hidden"
visibility_teilnehmer = "hidden"
print(basedir)
print(urkunden_file)
#________________________________________________________________________________________
temp_test = 'hi'
app.config.update(
    UPLOADED_PATH=os.path.join(basedir, 'files'),
    # Flask-Dropzone config:
    DROPZONE_MAX_FILE_SIZE=2024,  # maximale file größe
    DROPZONE_TIMEOUT=5 * 60 * 1000,  # maximale uploade dauer (hier 5 min)

)
dropzone = Dropzone(app)
#return render_template(HTML datei) ruft die angegebene website auf
#request.form.get(name des Buttons) == 'Wert des Buttons': prüft ob button gedrückt wurde
#die methode hinter @app.route wird aufgerufen aus der html aufgerufen action="/" in html legt fest welches @app.route
@app.route('/',methods=['GET', 'POST'])
def index():
    print("index")
    if flask.request.method == 'POST':
        if flask.request.form.get('einstellungen_button') == 'Einstellungen':
            print('einstellungen Klick')
            return flask.render_template('einstellungen.html', display_teilnehmer = "none", display_urkunde ="none", display_teilnehmer_list = "none")
            pass  # do something
        elif flask.request.form.get('auswertung_start') == 'Auswertung start':
            print('start Auswertung click')
            main.auswertung.auswertung(main.auswertung)
            return flask.render_template('home.html')
        elif flask.request.form.get("tn_hinzufügen") == "tn_hinzufügen":
            print("tn hinzufügen click")
            disziplin_liste = main.get_disziplinen()
            print(disziplin_liste)
            return flask.render_template('zeiten_bearbeiten.html', disziplin_list= disziplin_liste)
        elif flask.request.form.get('home') == 'home':
            disziplinen_list = main.get_disziplinen()
            return  flask.render_template('home.html')
        elif flask.request.form.get('zeitmessung_start') == 'zeitmessung start':
            disziplinen = main.get_disziplinen()
            print(disziplinen)
            return flask.render_template('zeit_stoppen.html',display_zeit_stoppen_tn="none",display_startup="True",disziplinen = disziplinen)
            print('config uploade')
        elif flask.request.form.get('export') =='export':
            print("export click")
            #test.run()
            main.export.export(main.export)
            liste = main.get_export_files()
            return flask.render_template('upload.html', files=liste)
        elif flask.request.form.get("downloade_start") == 'downloade files':
            liste = main.get_export_files()
            return flask.render_template('upload.html', files =liste)
        elif flask.request.form.get("new_user_bt") == 'Neuer Teilnehmer':
            return flask.render_template('new_tn.html',tn_num = "")
        else:
            pass  #
    elif flask.request.method == 'GET':
        disziplinen_list = main.get_disziplinen()
        return flask.render_template('home.html')
#_____________________________________________________________uploade Time
@app.route("/zeiten_bearbeiten", methods=['POST', 'GET'])
def zeiten_bearbeiten():
    disziplinen_list = main.get_disziplinen()

    return flask.render_template("zeiten_bearbeiten", disziplinen_list = disziplinen_list)
@app.route('/upload_time',methods=[ 'POST','GET'])
def upload_time():
    print('uploade time')
    if flask.request.method == 'POST':
        print('uploade time file')
        f = flask.request.files.get('file') #empfängt neuen file
        if f:
            file_path = Path(time_file, f.filename)
            # falls der file schon vorhanden ist wird dieser numeriert abgespeichert dafür ist die gesammte if
            if file_path.is_file():
                vorhanden = True
                x = 1
                filename = f.filename
               # print(filename)
                filename_list = filename.split('.')
                filename = filename_list[0]
                filename_endung = filename_list[1]
                while vorhanden:
                    file_path = Path(time_file, (filename + str(x) + '.' + filename_endung))
                    if file_path.is_file():
                        vorhanden = True
                    else:
                        vorhanden = False
            f.save(file_path)
            disziplinen_list = main.get_disziplinen()
            return flask.render_template("home.html",disziplinen=disziplinen_list)
        if flask.request.form.get('home') == 'home':
            disziplinen_list = main.get_disziplinen()
            return flask.render_template('home.html',disziplinen=disziplinen_list)
    disziplinen_list = main.get_disziplinen()
    return flask.render_template("home.html",disziplinen=disziplinen_list)
@app.route('/uploade_urkunde',methods=[ 'POST','GET'])
def uploade_urkunde():
    print('uploade urkunde')
    if flask.request.method == 'POST':
        f = flask.request.files.get('file')
        if f:
            print('get file')
            file_path =urkunden_file + '/' + f.filename
            f.save(file_path)
        return flask.render_template("einstellungen.html")
    return flask.render_template("einstellungen.html")
#______________________________________________uploade teilnehmer

@app.route('/zeit_messung',methods=[ 'POST','GET'])
def zeit_messung():
    print("zeit messung")
    if flask.request.method == 'POST':
        if flask.request.form.get('back_button') == 'back':
            Web_interface.temp_class.temp_teilnehmer_nummer = temp_class.temp_teilnehmer_nummer[0:len(temp_class.temp_teilnehmer_nummer) - 1]
            print('back button push')
            return flask.render_template('zeit_stoppen.html', prediction_text=str(temp_class.temp_teilnehmer_nummer),display_startup="none",display_zeit_stoppem_tn="True")
        elif flask.request.form.get('enter_button') == 'enter':
            print('enter button push')
            teilnehmer_zeit = main.stoppuhr.new_time(main.stoppuhr,temp_class.start_time,temp_class.disziplin,temp_class.temp_teilnehmer_nummer)
            teilenhmer_nummer = temp_class.temp_teilnehmer_nummer
            teilnehmer = main.get_teilnehmer_infos(teilenhmer_nummer)
            teilnehmer_zeit = main.auswertung.decode_time(main.auswertung,teilnehmer_zeit)
            print(teilnehmer['Nachname'])
            teilnehmer_name =  teilnehmer['Vorname'].values[0] + ' ' +  teilnehmer['Nachname'].values[0]
            teilnehmer = dict( nummer = teilenhmer_nummer,name = teilnehmer_name, zeit = teilnehmer_zeit)
            temp_class.last_teilnehmer_list.append(teilnehmer)
            temp_class.temp_teilnehmer_nummer = ''
            return flask.render_template('zeit_stoppen.html', prediction_text=' ',display_startup="none",display_zeit_stoppen_tn="True",display_zeit_stoppen_ohne_tn = "none",teilnehmer_list = temp_class.last_teilnehmer_list)
        elif flask.request.form.get('stopp_button') == 'stoppen':
            print("stopp button push")
            teilnehmer_zeit = main.stoppuhr.new_time(main.stoppuhr, temp_class.start_time, temp_class.disziplin)
            temp_class.zeiten_list.append(teilnehmer_zeit)
            return flask.render_template('zeit_stoppen.html', prediction_text=' ', display_startup="none", display_zeit_stoppen_ohne_tn="True",display_zeit_stoppen_tn="none",zeiten_list=temp_class.zeiten_list)
        elif flask.request.form.get('start_button') == 'start':
            disziplin = flask.request.form.get('disziplinselect')
            mit_tn = flask.request.form.get("mit_startnr")
            print(mit_tn)
            print(disziplin)
            temp_class.disziplin = disziplin
            temp_class.start_time = time.monotonic()
            print('hi')
            if mit_tn == "true":
                return flask.render_template('zeit_stoppen.html',display_zeit_stoppen_tn="True",display_startup="none",display_zeit_stoppen_ohne_tn ="none")
            else:
                return flask.render_template('zeit_stoppen.html',display_zeit_stoppen_tn="none",display_startup= "none",display_zeit_stoppen_ohne_tn ="True")
        elif flask.request.form.get('Zahlen_button'):
            temp_class.temp_teilnehmer_nummer  = temp_class.temp_teilnehmer_nummer + str(flask.request.form.get('Zahlen_button'))
            print(flask.request.form.get('Zahlen_button'))
            return flask.render_template('zeit_stoppen.html', prediction_text=str(temp_class.temp_teilnehmer_nummer ),display_zeit_stoppen_tn ="True",display_startup='none',teilnehmer_list = temp_class.last_teilnehmer_list)
        elif flask.request.form.get("home_button") == "home":
            disziplinen_list = main.get_disziplinen()
            return flask.render_template("home.html")

@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    print('downloade')
    return flask.send_from_directory(os.path.abspath(".") + '/files/export/', filename, as_attachment=True)
@app.route('/downloade_urkunden/<path:filename>', methods=['GET'])#sendet urkunde an client
def downloade(filename):
    print('downloade')
    return flask.send_from_directory(os.path.abspath(".") + '/files/Urkunden_Zusammenfassung/', filename, as_attachment=True)

@app.route('/einstellungen',methods=['POST','GET'])
def einstellungen():

    print('methode einstellungen')
    if flask.request.method == 'POST':
        if flask.request.form.get('urkunden_bt'):
            urkunden = main.get_urkunden_files()
            print(urkunden)
            return flask.render_template('einstellungen.html',display_urkunde="True",display_teilnehmer="none",urkunden=urkunden)
        elif flask.request.form.get('home')=='home':
            print('home website aufgerufen')
            disziplinen_list = main.get_disziplinen()
            return flask.render_template('home.html',disziplinen=disziplinen_list)
        elif flask.request.form.get("reset_button") == 'reset':
            main.reset()
            return  flask.render_template('einstellungen.html',uploade_file_display='block')
        elif flask.request.form.get('uploade_urkunde_bt'):
            print("uploade urkunde")
            neue_urkunde = flask.request.files['urkunde']
            disziplin = flask.request.form["disziplin_urkunde"]
            neue_urkunde.save( os.path.join(os.path.abspath(".") + "/files/Urkunden_Zusammenfassung/" + disziplin + ".docx"))
            urkunden = main.get_urkunden_files()
            print(urkunden)
            return flask.render_template('einstellungen.html', display_urkunde="True", display_teilnehmer="none",urkunden=urkunden)
        elif flask.request.form.get('teilnehmer_bt'):
            print("teilnehmer bt klick")
            return flask.render_template('einstellungen.html', display_urkunde ="none", display_teilnehmer="True", display_teilnehmer_list = "none")
        elif flask.request.form.get("show_teilnehmer_bt"):#
            print("teilnehmer list bt klick")
            teilnehmer_list = main.get_teilnehmer_list()
            return flask.render_template('einstellungen.html', display_urkunde="none",display_teilnehmer="True",diplay_teilnehmer_list = "True",Teilnehmer_array = teilnehmer_list)
        elif flask.request.form.get("uploade_teilnehmer_bt"):
            print('uploade teilnehmer bt klicks')
            neuer_teilnehmer_file = flask.request.files['teilnehmer_file']
            print(190)
            neuer_teilnehmer_file.save(os.path.join(os.path.abspath(".") + "/files/Teilnehmer.xlsx"))
            print(192)
            main.new_teilnehmer_file()
            return flask.render_template('einstellungen.html', display_urkunde ="none", display_teilnehmer="True",display_teilnehmer_list="none")

@app.route('/new_tn', methods=['POST','GET'])
def new_tn():
    print("new tn")
    if flask.request.method == 'POST':
        if flask.request.form.get('save_bt'):
            vorname = flask.request.form['Vorname_textfield']
            nachname = flask.request.form["Nachname_textfield"]
            ak= flask.request.form["AK_textfield"]
            disziplin = flask.request.form["Disziplin_textfield"]
            verein = flask.request.form["Verein_textfield"]
            tn_num = main.new_tn(vorname,nachname,ak,disziplin,verein)
            return flask.render_template('new_tn.html', tn_num =tn_num)
    return flask.render_template("new_tn.html",tn_num="")
@app.route('/files',methods=['POST','GET'])
def files():
    print('files')
    if flask.request.method == 'POST':
        if flask.request.form.get('löschen_button'):
            for file in flask.request.form.getlist('file_checkbox'):
                print('lösche: ' + file )
    return flask.render_template('einstellungen.html')
class speicher:
    def __init__(self):
        self.temp_teilnehmer_nummer = ''
        self.last_teilnehmer_list = []
        self.start_time = ''
        self.disziplin = ''
        self.teilnehmer_list = []
        self.zeiten_list=[]

uhr = flask.Flask(__name__, static_folder='static')

def start_Web_interface():
    #app.run(host='DESKTOP-91HA56Q', port='5000')
    app.run()
temp_class = speicher()

