# -*- coding: utf-8 -*-
# 1. Import von 6 Bibliotheken
import requests
import datetime
import csv
import xlsxwriter
import sys
from bs4 import BeautifulSoup

# 2. Definition von 5 Funktionen
def Keyword_Liste():
    keywords_datei = open('Keywords.csv', encoding='utf-8-sig', errors='ignore')
    keywords_elemente = list(csv.reader(keywords_datei))
    keyword_mensch_liste = [keywords_elemente[element][0] for element in range(len(keywords_elemente))]
    keyword_maschine_liste = [keywords_elemente[element][0].replace(" ","+") for element in range(len(keywords_elemente))]
    return keyword_mensch_liste, keyword_maschine_liste

def URL_Liste(keyword_maschine_liste):
    url_liste=['https://www.google.de/search?q=' + str(keyword) + '&num=100' for keyword in keyword_maschine_liste] 
    return url_liste

def Google_Suche(URL_liste):
    aktuelles_keyword = 1
    fake_browser = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:77.0) Gecko/20100101 Firefox/77.0'}
    link_liste=[]
    ergebnis_liste=[]
    
    for URL in URL_liste:
        print(">>Suche Keywords (Keyword ", aktuelles_keyword, "/", len(URL_liste), ")", sep='', end='\r')
        
        req_google= requests.get(URL, headers=fake_browser).content
        soup_google= BeautifulSoup(req_google, 'html.parser')
        
        pfad_liste=[pfad['href'] for pfad in soup_google.select('div[class="r"]>a')]
        
        if any('your-website.de' in pfad for pfad in pfad_liste):
            for pfadnummer in range(len(pfad_liste)):
                if pfad_liste[pfadnummer][0:28] == 'https://www.your-website.de/':
                    link_liste.append(pfad_liste[pfadnummer])
                    
                    if soup_google.select('h2:contains("Hervorgehobenes Snippet aus dem Web")'):
                        ergebnis=pfadnummer
                    else:
                        ergebnis=pfadnummer+1
                        
                    ergebnis_liste.append(ergebnis)
                    aktuelles_keyword +=1
                    break
        else:
            ergebnis_liste.append("Nicht in den ersten 100 Google-Ergebnissen")
            link_liste.append("")
            aktuelles_keyword +=1
            
    return ergebnis_liste, link_liste

def Datum():
    aktuelles_datum = datetime.datetime.now()
    datum_maschine= aktuelles_datum.strftime("%Y-%m-%d-%H-%M")
    datum_mensch= aktuelles_datum.strftime("%d.%m.%Y")              
    
    return datum_maschine, datum_mensch

def Excel_Datei(keyword_mensch_liste, ergebnis_liste, link_liste, datum_maschine, datum_mensch):
    excel_dateiname = 'Keywords_' + datum_maschine + '.xlsx'
    excel_datei = xlsxwriter.Workbook(excel_dateiname)
    excel_arbeitsblatt = excel_datei.add_worksheet('Keywordplatzierung')

    excel_ueberschriften = ["Keyword","URL","Google-Platzierung am " + datum_mensch]
    fett = excel_datei.add_format({'bold': 1})
	
    excel_arbeitsblatt.write_row('A1', excel_ueberschriften, fett)
    excel_arbeitsblatt.write_column('A2', keyword_mensch_liste)
    excel_arbeitsblatt.write_column('B2', link_liste)
    excel_arbeitsblatt.write_column('C2', ergebnis_liste)
	
    excel_datei.close()
    print("\n>>Datei " + excel_dateiname + " erstellt")
    input("Enter drücken, um das Programm zu verlassen")
    sys.exit()
    
# 3. Hauptprogramm: Ausführung von fünf Funktionen im Hauptprogramm, um die passenden Variablen zu füllen

keyword_mensch_liste, keyword_maschine_liste = Keyword_Liste()
URL_liste = URL_Liste(keyword_maschine_liste)                          
ergebnis_liste, link_liste = Google_Suche(URL_liste)
datum_maschine, datum_mensch = Datum()
Excel_Datei(keyword_mensch_liste , ergebnis_liste, link_liste, datum_maschine, datum_mensch)