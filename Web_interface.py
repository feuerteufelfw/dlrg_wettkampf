#import gedöns
import os
from pathlib import Path
import flask
import main
from flask_dropzone import Dropzone
import Web_interface
import time
#____________________________________________________definition aller verwendeten pfade
basedir = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'files')
app = flask.Flask(__name__,static_folder='static')
time_file = os.path.join(basedir, 'time')
config_file = os.path.join(basedir, 'config')
teilnehmer_file = basedir + '/files'
urkunden_file = basedir  + "/files/Urkunden_Zusammenfassung"
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
            return flask.render_template('einstellungen.html')
            pass  # do something
        elif flask.request.form.get('auswertung_start') == 'Auswertung start':
            print('start Auswertung click')
            main.auswertung(temp_class.disziplin)
            liste= []
            liste.append('Urkunden_Gesamt.pdf')
            liste.append('hi.pdf')
            return flask.render_template('upload.html',files=liste)
        elif flask.request.form.get('home') == 'home':
            return  flask.render_template('home.html')
        elif flask.request.form.get('zeitmessung_start') == 'zeitmessung start':
            disziplinen_list = ["2000m", "500m", "1000m"]
            return flask.render_template('zeit_stoppen.html',visibility="hidden",visibility_startup="visible",disziplinen=disziplinen_list)
            print('config uploade')
        else:
            pass  #
    elif flask.request.method == 'GET':
        return flask.render_template('home.html')
#_____________________________________________________________uploade Time
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
            return flask.render_template("home.html")
        if flask.request.form.get('home') == 'home':
            return flask.render_template('home.html')
    return flask.render_template("home.html")
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
@app.route('/upload_teilnehmer',methods=[ 'POST','GET'])
def upload_teilnehmer():
    print('uploade teilnehmer')
    if flask.request.method == 'POST':
        print('uploade teilnehmer file')
        f = flask.request.files.get('file')
        if f:
            file_path = teilnehmer_file + '/' + f.filename
            f.save(file_path)
            main.new_teilnehmer()
            return flask.render_template("einstellungen.html")
        if flask.request.form.get('home') == 'home':
            return flask.render_template('home.html')
        return flask.render_template("einstellungen.html")
@app.route('/zeit_messung',methods=[ 'POST','GET'])
def zeit_messung():
    print("zeit messung")
    if flask.request.method == 'POST':
        if flask.request.form.get('back_button') == 'back':
            Web_interface.temp_class.temp_teilnehmer_nummer = temp_class.temp_teilnehmer_nummer[0:len(temp_class.temp_teilnehmer_nummer) - 1]
            print('back button push')
            return flask.render_template('zeit_stoppen.html', prediction_text=str(temp_class.temp_teilnehmer_nummer),visibility_startup="hidden",visibility="visible",disziplin=disziplin)
        elif flask.request.form.get('enter_button') == 'enter':
            print('enter button push')
            temp_class.last_teilnehmer_list.append(temp_class.temp_teilnehmer_nummer)
            main.stoppuhr.new_time(main.stoppuhr,temp_class.temp_teilnehmer_nummer,temp_class.start_time)
            temp_class.temp_teilnehmer_nummer = ''
            return flask.render_template('zeit_stoppen.html', prediction_text=' ',visibility_startup="hidden",visibility="visible",disziplin=disziplin)
        elif flask.request.form.get('start_button') == 'start':
            speicher.disziplin = flask.request.form["Disziplin"]
            print(speicher.disziplin)
            if speicher.disziplin:
                temp_class.disziplin = speicher.disziplin
                temp_class.start_time = time.monotonic()
                print('hi')
                return flask.render_template('zeit_stoppen.html',visibility="visible",visibility_startup="hidden",disziplin=speicher.disziplin)
            else:
                return flask.render_template('zeit_stoppen.html',visibility = "hidden",visibillity_startup = 'visible',disziplin=speicher.disziplin)
        elif flask.request.form.get('Zahlen_button'):
            temp_class.temp_teilnehmer_nummer  = temp_class.temp_teilnehmer_nummer + str(flask.request.form.get('Zahlen_button'))
            print(flask.request.form.get('Zahlen_button'))
            return flask.render_template('zeit_stoppen.html', prediction_text=str(temp_class.temp_teilnehmer_nummer ),visibility ="visible",visibility_startup='hidden',disziplin=speicher.disziplin)
        elif flask.request.form.get("home_button") == "home":
            return flask.render_template("home.html")

@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    print('downloade')
    return flask.send_from_directory('C:\\Users\\feuer\\OneDrive\\Dokumente\\DLRG\\Wettkampfrechner\\', filename, as_attachment=True)
@app.route('/einstellungen',methods=['POST','GET'])
def einstellungen():
    print('methode einstellungen')
    if flask.request.method == 'POST':
        if flask.request.form.get('neues_datum') == 'neues Datum':
            datum = flask.request.form('datum_textfield')
            print("Datum: " + datum)
        elif flask.request.form.get('uploade_button'):
            return flask.render_template('einstellungen.html',uploade_file_display='block')
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

def start_Web_interface():
    app.run()
temp_class = speicher()

