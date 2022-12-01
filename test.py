from datetime import datetime

datum = datetime.today()
tag = datum.day
month = datum.month
year = datum.year
print('Tag: ' + str(tag))
print('Monat: ' + str(month))
print('Jahr: ' + str(year))
datum = str(tag) + '.' + str(month) + '.' + str(year)
print(datum)