from subprocess import  Popen
LIBRE_OFFICE = r"/snap/bin/libreoffice.writer"

def convert_to_pdf(input_docx, out_folder):
    p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_docx])
    print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
    p.communicate()


sample_doc = '/home/fabian/Schreibtisch/Wettkampfrechner/temp/dokument.docx'
out_folder = '/home/fabian/Schreibtisch/Wettkampfrechner/temp/'
convert_to_pdf(sample_doc, out_folder)