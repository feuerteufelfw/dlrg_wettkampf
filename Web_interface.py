#import gedöns
import os
from pathlib import Path
from flask import *
import main
from flask_dropzone import Dropzone
import Web_interface
import time
#____________________________________________________definition aller verwendeten pfade
basedir = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
app = Flask(__name__,template_folder='tamplate')
time_file = os.path.join(basedir, 'time')
config_file = os.path.join(basedir, 'config')
teilnehmer_file = os.path.join(basedir, 'teilnehmer')
urkunden_file = UPLOAD_FOLDER,'\\Urkunden\\'
#________________________________________________________________________________________
temp_test = 'hi'
app.config.update(
    UPLOADED_PATH=os.path.join(basedir, 'uploade'),
    # Flask-Dropzone config:
    DROPZONE_MAX_FILE_SIZE=2024,  # maximale file größe
    DROPZONE_TIMEOUT=5 * 60 * 1000,  # maximale uploade dauer (hier 5 min)
    DROPZONE_UPLOAD_ACTION='upload',
    DROPZONE_UPLOAD_MULTIPLE=True,#erlaubt parallele uploads
    DROPZONE_PARALLEL_UPLOADS=3#erlaubt bis zu drei uploads gleichzeitig
)
app.static_folder = 'static'
dropzone = Dropzone(app)
#return render_template(HTML datei) ruft die angegebene website auf
#request.form.get(name des Buttons) == 'Wert des Buttons': prüft ob button gedrückt wurde
#die methode hinter @app.route wird aufgerufen aus der html aufgerufen action="/" in html legt fest welches @app.route
@app.route('/',methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if request.form.get('einstellungen_button') == 'Einstellungen':
            print('einstellungen Klick')
            return render_template('einstellungen.html')
            pass  # do something
        elif request.form.get('auswertung_start') == 'Auswertung start':
            print('start Auswertung click')
            main.auswertung(temp_class.disziplin)
            liste= []
            liste.append('Urkunden_Gesamt.pdf')
            liste.append('hi.pdf')
            return render_template('upload.html',files=liste)
        elif request.form.get('home') == 'home':
            return  render_template('home.html')
        elif request.form.get('zeitmessung_start') == 'zeitmessung start':
            return render_template('zeit_stoppen.html',visibility="hidden",visibility_startup="visible")
            print('config uploade')
        else:
            pass  #
    elif request.method == 'GET':
        return render_template('home.html')
#_____________________________________________________________uploade Time
@app.route('/upload_time',methods=[ 'POST','GET'])
def upload_time():
    if request.method == 'POST':
        print('uploade time file')
        f = request.files.get('file') #empfängt neuen file
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
            return render_template("home.html")
        if request.form.get('home') == 'home':
            return render_template('home.html')
    return render_template("home.html")
@app.route('/upload_urkunde',methods=[ 'POST','GET'])
def upload_urkunde():
    if request.method == 'POST':
        if request.form('disziplin_textfield'):
            speicher.temp_disziplin = request.form('disziplin_textfield')
            print(speicher.temp_disziplin)
        else:
            print('uploade urkunde file')
            speicher.temp_disziplin = request.form('disziplin_textfield')
            print(speicher.temp_disziplin)
            f = request.files.get('file')
            print(f)
            if f:
                print('get file')
                file_path = Path(urkunden_file,'Urkunde_', speicher.temp_disziplin,'.docx')
                f.save(file_path)
            return render_template("einstellungen.html")

    return render_template("home.html")

#______________________________________________uploade teilnehmer
@app.route('/upload_teilnehmer',methods=[ 'POST','GET'])
def upload_teilnehmer():
    if request.method == 'POST':
        print('uploade teilnehmer file')
        f = request.files.get('file')
        if f:
            file_path = Path(teilnehmer_file, f.filename)
            f.save(file_path)
            main.new_teilnehmer()
            return render_template("home.html")
        if request.form.get('home') == 'home':
            return render_template('home.html')
        return render_template("home.html")
@app.route('/zeit_messung',methods=[ 'POST','GET'])
def zeit_messung():
    if request.method == 'POST':
        if request.form.get('back_button') == 'back':
            Web_interface.temp_class.temp_teilnehmer_nummer = temp_class.temp_teilnehmer_nummer[0:len(temp_class.temp_teilnehmer_nummer) - 1]
            print('back button push')
            return render_template('zeit_stoppen.html', prediction_text=str(temp_class.temp_teilnehmer_nummer))
        elif request.form.get('enter_button') == 'enter':
            print('enter button push')
            temp_class.last_teilnehmer_list.append(temp_class.temp_teilnehmer_nummer)
            main.stoppuhr.new_time(main.stoppuhr,temp_class.temp_teilnehmer_nummer,temp_class.start_time)
            temp_class.temp_teilnehmer_nummer = ''
            return render_template('zeit_stoppen.html', prediction_text=' ')
        elif request.form.get('start_button') == 'start':
            disziplin = request.form["disziplin_textfield"]
            if disziplin:
                temp_class.disziplin = disziplin
                temp_class.start_time = time.monotonic()
                print('hi')
                return render_template('zeit_stoppen.html',visibility="visible",visibility_startup="hidden")
            else:
                return render_template('zeit_stoppen.html',visibility = "hidden",visibillity_startup = 'visible')
        elif request.form.get('Zahlen_button'):
            temp_class.temp_teilnehmer_nummer  = temp_class.temp_teilnehmer_nummer + str(request.form.get('Zahlen_button'))
            print(request.form.get('Zahlen_button'))
            return render_template('zeit_stoppen.html', prediction_text=str(temp_class.temp_teilnehmer_nummer ),visibility ="visible",visibility_startip='hidden')
        elif request.form.get("home_button") == "home":
            return render_template("home.html")

@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    return send_from_directory('C:\\Users\\feuer\\OneDrive\\Dokumente\\DLRG\\Wettkampfrechner\\', filename, as_attachment=True)
@app.route('/einstellungen',methods=['POST','GET'])
def einstellungen():
    print('methode einstellungen')
    if request.method == 'POST':
        if request.form.get('neues_datum') == 'neues Datum':
            datum = request.form('datum_textfield')
            print("Datum: " + datum)
        elif request.form.get('uploade_button'):
            return render_template('einstellungen.html',uploade_file_display='block')
@app.route('/files',methods=['POST','GET'])
def files():
    print('hi')
    if request.method == 'POST':
        if request.form.get('löschen_button'):
            for file in request.form.getlist('file_checkbox'):
                print('lösche: ' + file )
    return render_template('einstellungen.html')
class speicher:
    def __init__(self):
        self.temp_teilnehmer_nummer = ''
        self.last_teilnehmer_list = []
        self.start_time = ''
        self.disziplin = ''
        self.temp_disziplin =""

def start_Web_interface():
    app.run()
temp_class = speicher()

