from docx2pdf import convert
import pythoncom
def run():
    pythoncom.CoInitialize()
    convert("C://Users//Dell Latitude//PycharmProjects//dlrg_wettkampf//files//temp//docx//dokument4.docx")
