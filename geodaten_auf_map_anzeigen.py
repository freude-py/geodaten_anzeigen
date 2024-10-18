# Geodaten anzeigen auf einer map
# die Geodaten werden aus einer Excel Tabelle ausgelesen
# Datenstruktur: PLZ, ORT, geo1, geo2, checker (0=noch nicht gecheckt, 1=passt genau, 2=passt fast, 3=passt nicht)
# 1. Zeile = Header

import tkinter  as tk 
from tkinter import *
import webbrowser
import folium
from folium.plugins import MarkerCluster
from openpyxl import load_workbook

excelfile = "C:\GW_Daten\#99 Privat\Programmierung\Python\geodaten\geodaten_auf_map_anzeigen.xlsx"
register_geodaten = "geodaten"
output_map = "C:\GW_Daten\#99 Privat\Programmierung\Python\geodaten\geodaten_auf_map.html"

def map_anzeigen():

    # Mappe initialisieren
    mapObj = folium.Map(zoom_start=12, location=[47.2657315, 11.3939171], control_scale=True, min_zoom=7, width=1800, height=950,)
    mc1 = MarkerCluster(name="check offen").add_to(mapObj)              # Create a MarkerCluster group
    mc2 = MarkerCluster(name="passt genau").add_to(mapObj)              # Create a MarkerCluster group
    fg1 = folium.FeatureGroup(name="passt fast").add_to(mapObj)        # Create a FeatureGroup
    fg2 = folium.FeatureGroup(name="passt nicht").add_to(mapObj)        # Create a FeatureGroup
    folium.FitOverlays().add_to(mapObj)
    folium.LayerControl().add_to(mapObj)

    # geodaten aus Excel auslese und Punkte auf Karte setzen
    book = load_workbook(excelfile)
    sheet = book[register_geodaten]

    a=0
    for row in sheet.rows:  # Die Datensätze aus Excel auslesen
        
        if a==0: # Erste Zeile überlesen weil Header 
            a=1
            continue
    
        ort = row[1].value    # Ort
        geo1 = row[2].value   # geo1
        geo2 = row[3].value   # geo2
        checker = row[4].value # 0=noch nicht gecheckt, 1=passt genau, 2=passt fast, 3=passt nicht
        #print(ort, geo1, geo2)
         
        # Marker in map lt. geodaten einfügen
        if checker == 0:
            folium.Marker(location=[geo1, geo2],
                tooltip =(ort),                           # hier kann man Infos für den tooltip eingeben
                popup=folium.Popup(ort, max_width=500),   # hier kann man Infos für die popup anzeigen eingeben
                icon=folium.Icon(icon='glyphicon-star', color='white'),
            ).add_to(mc1)

        if checker == 1:      # eigener Layer für eine bestimmte Datenart
            folium.Marker(location=[geo1, geo2],
                tooltip=(ort), 
                popup=folium.Popup(ort, max_width=500),
                icon=folium.Icon(icon='glyphicon-star', color='green'),
            ).add_to(mc2)
            
        if checker == 2:      # eigener Layer für eine bestimmte Datenart
            folium.Marker(location=[geo1, geo2],
                tooltip=(ort), 
                popup=folium.Popup(ort, max_width=500),
                icon=folium.Icon(icon='glyphicon-star', color='blue'),
            ).add_to(fg1)

        if checker == 3:      # eigener Layer für eine bestimmte Datenart
            folium.Marker(location=[geo1, geo2],
                tooltip=(ort), 
                popup=folium.Popup(ort, max_width=500),
                icon=folium.Icon(icon='glyphicon-star', color='red'),
            ).add_to(fg2)

    mapObj.save(output_map)
    webbrowser.open(output_map)

    print('Programm fehlerfrei druchlaufen')

    
if __name__=='__main__':
    map_anzeigen()

