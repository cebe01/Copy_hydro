# -*- coding: utf-8 -*-
"""
Spyder Editor

Skript zum Herauslesen von Wasserstandsdaten aus verschiedenen Excelfiles
Bei Fragen: Cb
"""

import pandas as pd
import numpy as np
import os 

path = 'X:\Cb\Hydrodaten_Cb' #Pfad zu den Files, Files müssen im Ordner nach Datum sortiert sein
sheet = 'Wasserzähler I' #Excelsheet, wo die Station ist
station = 'GWPW Schmalholz' #name der Station, sowie sie im Excelfile in der Zeile 23 steht



files = [f for f in os.listdir(path)]# Alle Excel-Files im Ordner zusammensuchen

datum = list() #List für Datumsdaten initialisieren
wasserstand = list() #List für Wasserstandsdaten initialisieren


for file in files: #Gehe durch alle Excel-Files
    
    try: df = pd.read_excel(path+'/'+file, sheet_name=sheet) #Try Load des Excelfiles (Fehlermeldung, welche durch Nicht-Excelfiles)
    except: continue
    
    for i in range(0,999): #Dieser Loop sucht die Spalte, wo die Station gespeichert ist
        if df.iloc[22,i] == station:
            station_colum = i
            break #Loop wird gestoppt, wenn Station gefunden
        else: continue
    

    for i in range (29,29+31): #Daten befinden sich ab Zeile 29 im excel
        if pd.isnull(df.iloc[i,2]) == False: #Mit dieser Bedingung werden die verschiedenen Anzahl Tage pro Monat berücksichtigt
    
          datum.append(df.iloc[i,2]) #Datum stehen in der Spalte 2
          wasserstand.append(df.iloc[i,station_colum]) #Wasserstände kopieren
      
        else: break
        

tosave = pd.DataFrame(({'datum': datum, 'wasserstand': wasserstand})) #Listen zu DataFrame zusammensetzen

if not os.path.exists(path+ '\Auswertungen'): #Schauen, ob Ordner bereits existiert
    os.makedirs(path+ '\Auswertungen') #Ordner erstellen
    
tosave.to_excel(path+ '\Auswertungen'+'\ges_hydrodaten_'+station+'.xlsx') #Gesamelte Daten als Excel speichern