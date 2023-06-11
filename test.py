from docx2pdf import convert
import pythoncom
import datetime

geburtsdatum = datetime.date(2003,6,11)
datum = datetime.datetime.today().date()
alter = datum.year - geburtsdatum.year
if datum.month < geburtsdatum.month:
    alter = alter -1
elif datum.month == geburtsdatum.month:
  if datum.day < geburtsdatum.day:
    alter = alter-1

print(alter)
def run():
    pythoncom.CoInitialize()
    convert("C://Users//Dell Latitude//PycharmProjects//dlrg_wettkampf//files//temp//docx//dokument4.docx")
