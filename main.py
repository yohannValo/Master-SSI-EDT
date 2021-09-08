import pandas as pd
from ics import Calendar, Event
import os

c = Calendar()
xl = pd.read_excel("/home/yohann/PycharmProjects/edt/EduTemps21.xlsx", usecols="B:H,N,O")
for i in range(len(xl)):
    for j in range(1, 7):
        if isinstance(xl.iloc[i][j], str) and len(xl.iloc[i][j]) > 1:

            e = Event()
            e.name = xl.iloc[i][j]
            if j < 3:
                e.location = str(xl.iloc[i][7])
            else:
                e.location = str(xl.iloc[i][8])

            dateDebut = xl.iloc[i][0]
            if j == 1:
                dateDebut = dateDebut.replace(hour=8)
            elif j == 2:
                dateDebut = dateDebut.replace(hour=9)
            elif j == 3:
                dateDebut = dateDebut.replace(hour=10)
            elif j == 4:
                dateDebut = dateDebut.replace(hour=12)
            elif j == 5:
                dateDebut = dateDebut.replace(hour=14)
            elif j == 6:
                dateDebut = dateDebut.replace(hour=16)

            e.begin = dateDebut

            dateFin = xl.iloc[i][0]

            if j == 1:
                dateFin = dateFin.replace(hour=9)
            elif j == 2:
                dateFin = dateFin.replace(hour=10)
            elif j == 3:
                dateFin = dateFin.replace(hour=12)
            elif j == 4:
                dateFin = dateFin.replace(hour=14)
            elif j == 5:
                dateFin = dateFin.replace(hour=16)
            elif j == 6:
                dateFin = dateFin.replace(hour=18)
            e.end = dateFin
            c.events.add(e)

with open('my.ics', 'w') as f:
    f.write(str(c))

cmd = 'sed -i "s/0000Z/0000/g" my.ics'

os.system(cmd)
