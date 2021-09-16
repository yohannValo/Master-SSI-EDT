import pandas as pd
from ics import Calendar, Event
from pytz import timezone
import os

tz = timezone("Europe/Paris")
c = Calendar() # creation d'un calendrier
xl = pd.read_excel("EduTemps21.xlsx", usecols="B:H,N,O") # on garde les colonnes de date matiere et salles de classe
for i in range(len(xl)):
    for j in range(1, 7):
        if isinstance(xl.iloc[i][j], str) and len(xl.iloc[i][j]) > 1: # on parcourt le dataframe en excluant les valeurs nulles ou les espaces simples

            e = Event() # on creer un event
            e.name = xl.iloc[i][j] #on lui donne le nom de la matiere

            if j < 3:
                e.location = str(xl.iloc[i][7]) #ici le salles sont repartis en fonction du matin et de l apres midi donc
            elif j > 3:                              #avant midi on se situera sur la colonne 7 et apres midi sur la 8
                e.location = str(xl.iloc[i][8])

            dateDebut = xl.iloc[i][0] # ce sont les dates du jours au format annee-mois-jour 00h00 on va donc changer l'heure
            dateFin = xl.iloc[i][0]

            if j == 1:# groupe de condition un peu degueu qui vont changer l'heure de fin de la date pour s'accorder avec l excel
                dateDebut = dateDebut.replace(hour=8, tzinfo=tz)
                dateFin = dateFin.replace(hour=9, tzinfo=tz)
            elif j == 2:
                dateDebut = dateDebut.replace(hour=9, tzinfo=tz)
                dateFin = dateFin.replace(hour=10, tzinfo=tz)
            elif j == 3:
                dateDebut = dateDebut.replace(hour=10, tzinfo=tz)
                dateFin = dateFin.replace(hour=12, tzinfo=tz)
            elif j == 4:
                dateDebut = dateDebut.replace(hour=12, tzinfo=tz)
                dateFin = dateFin.replace(hour=14, tzinfo=tz)
            elif j == 5:
                dateDebut = dateDebut.replace(hour=14, tzinfo=tz)
                dateFin = dateFin.replace(hour=16, tzinfo=tz)
            elif j == 6:
                dateDebut = dateDebut.replace(hour=16, tzinfo=tz)
                dateFin = dateFin.replace(hour=18, tzinfo=tz)

            e.begin = dateDebut # on affecte les dates a l 'event
            e.end = dateFin
            c.events.add(e) # on ajoute enfin l'evenement au calendrier

with open('my.ics', 'w') as f:
    f.write(str(c)) #on creer un fichier ics ou on écrit le calendrier en format ics
