import re
import json
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy import URL
import mysql.connector
from pathlib import Path
import csv
import pymysql
import time
import xlwt
import openpyxl
from openpyxl import Workbook

df=pd.DataFrame()
# Listen der Präfixe, Zirkumfixe und Suffixe
prefixes = [
    "a", "an", "ar", "ab", "aber", "after", "anti", "auf", "aus", "auto", "be", "bei", "bio", "de", "des", "dis",
    "durch", "ein", "emp", "ent", "entgegen", "er", "erz", "ex", "fehl", "fest", "fort", "ge", "gegen", "geo",
    "graf", "her", "herunter", "hin", "hinter", "hyper", "ident", "in", "im", "il", "ir", "inne", "inter", "ko", "kol",
    "kom", "kon", "kor", "kund", "los", "makro", "maxi", "mega", "mikro", "mini", "miss", "mit", "mono", "multi",
    "nach", "nano", "naut", "neo", "non", "para", "pflichtig", "phil", "phob", "poly", "post", "prä", "pro", "proto",
    "pseudo", "quasi", "re", "riesen", "rück", "schwieger", "semi", "stereo", "stief", "tele", "therm", "thermo",
    "trans", "ultra", "um", "un", "unter", "ur", "ver", "vize", "vor", "weg", "wett", "wider", "zer", "zu", "zurecht",
    "zurück", "zusammen", "zuwider", "zwischen"
]

circumfixes = [
    "be...ig", "be...t", "ge...e", "ge...ig", "ge...sel", "ge...t", "ver...ig"
]

suffixes = [
    "a", "abel", "ibel", "ade", "iade", "age", "aholic", "oholic", "oholiker", "aille", "al", "ell", "ament",
    "ement", "an", "and", "ant", "ent", "ante", "ente", "anz", "enz", "ar", "är", "arium", "arm", "artig", "ast", "at",
    "ee", "ei", "eierei", "el", "elchen", "erchen", "elle", "ens", "er", "erich", "erie", "ern", "esk", "ess", "esse",
    "isse", "ette", "eur", "euse", "fach", "fähig", "bold", "chen", "dings", "drom", "e", "i", "ian", "jan", "ice",
    "icht", "ie", "ier", "ieren", "ifizier", "isier", "iere", "ig", "ik", "iker", "ine", "ing", "ingen", "en", "ern",
    "er", "e", "ell", "ion", "tion", "ation", "ismus", "asmus", "ist", "it", "ität", "itis", "iv", "ativ", "ke", "lei",
    "lein", "lekt", "ler", "lich", "ling", "lings", "lon", "los", "mals", "maßen", "mäßig", "mini", "n", "nis", "o",
    "oid", "ol", "or", "ator", "itor", "os", "ös", "ose", "ow", "pflichtig", "reich", "rich", "sal", "sam", "schaft",
    "sche", "seitig", "sel", "sen", "skop", "tex", "thek", "trächtig", "tum", "ung", "ur", "voll", "wang", "wangen",
    "wart", "wärts", "weg", "weise", "werk", "wesen", "zid"
]

diminutive_suffixes = [
    'elchen', 'chen', 'lein', 'erl', 'al', 'el', 'rl', 'ele', 'elein', 'ale', 'i'
    ]

verb_diminutive_suffixes =[

]

umlautungen_und_ablautungen = {
    'a': ['ä', 'i', 'e', 'u', 'ie'],
    'e': ['a', 'o', 'ie', 'u'],
    'i': ['a', 'o', 'ie', 'u'],
    'o': ['e', 'ö', 'a', 'ie'],
    'u': ['o', 'ü', 'ie'],
    'au': ['äu', 'o', 'ie'],
    'ei': ['i', 'ie'],
    'ie': ['o'],
    'ö': ['o']
}

doppelkonsonanten = {
    'b': ['bb'],
    'd': ['dd'],
    'g': ['gg'],
    'k': ['kk'],
    'l': ['ll'],
    'm': ['mm'],
    'n': ['nn'],
    'p': ['pp'],
    'r': ['rr'],
    's': ['ss'],
    't': ['tt']
}


def read_data( name:str = 'tbl_eintraege.tab')->pd.DataFrame:
    """
    Diese Methode nimmt eine Eingabe und liest eine Datei mit diesem Namen aus dem gleichen Verzeichnis und gibt einen darauf basierenden Pandas Dataframe zurück
    """
    # Pfad zur CSV-Datei
    #hier relativen Pfad statt absoluten benutzen, dann muss lediglich die Datenbank im gleichen Ordner liegen
    csv_file = Path(__file__).parent.joinpath(name)

    rows = []

    with open(csv_file, encoding='utf-8') as file:
        for line in file:
            rows.append(line.strip().split('\t')[0:6])
    # DataFrame erstellen
    df = pd.DataFrame(rows[1:], columns=rows[0])
    return df


def find_ableitungen(data:pd.DataFrame)->dict:  
    """
    Funktionsweise:
        Erhalte einen pandas Dataframe, welcher mindestens die Spalten Grammatik, Grundform und Lemma enthalten muss. Finde alle Ableitungen der Grundformen des Dataframe und ordne diese in einem dict dem letzten Lemma zu.
        Dieses dict wird zurückgegeben. Es werden Explizit keine Diminuitiva erkannt.
    Eingabe:
        data:pd.DataFrame = dies ist ein pandas Dataframe, welcher die Daten enthalten soll. Diese Daten müssen mindestens Grundform, Lemma und Grammatik beinhalten.
    Ausgabe:
        Ableitungen:dict = dieses dictionary enthält alle Ableitungen des Dataframe data und ordnet diese dem letzten Lemma der Grundform zu.
    """
    Ableitungen=dict()
    data=data[['Grundform','Lemma']].drop_duplicates().groupby('Grundform')    #Dies generiert einen Dataframe, welcher zu jeder unterschiedlichen Grundform, alle dazugehörigen Lemma anzeigt
    for grundform, group in data:
        lemmata=group['Lemma'].tolist()
        ableitung,lemma = is_ableitung(grundform,lemmata) #Hier findet die Prüfung statt, ob es wirklich eine Ableitung ist, oder lediglich eine Grundform
        if ableitung:   
            if lemma in Ableitungen:
                Ableitungen[lemma].append(grundform)
            else:
                Ableitungen[lemma]=[grundform]
    return Ableitungen


def is_ableitung(grundform:str,lemmata:list)->tuple:
    """
    Funktionsweise:
        hier wird überprüft ob ein Wort mit gegebenen Lemmata eine Ableitung ist und gibt das letzte Lemma des Wortes zurück. Diminuitiva werden hier explizit ausgelassen.
    Eingabe:
        grundform:str= die Grundform des Wortes
        lemmata:list= Eine Liste mit allen Lemmata des Wortes
    Ausgabe:
        is_ableitung:boolean = Gibt an ob das Wort eine Ableitung ist
        last_lemma:str = das letzte Lemma des Wortes
    """
    last_lemma = find_last_lemma(grundform,lemmata)
    grundform = grundform.lower()
    lemmata = [lemma.lower() for lemma in lemmata]
    lemmata = [lemma for lemma in lemmata if lemma not in prefixes]
    is_ableitung=False
    for prefix in prefixes:
        if grundform.startswith(prefix):
            if all(not ((x in grundform[len(prefix):]) ^ (x in grundform)) for x in lemmata):
                is_ableitung = True
                break
    # Prüfen auf Zirkumfixe
    if not is_ableitung:
        for circumfix in circumfixes:
            parts = circumfix.split("...")
            if grundform.startswith(parts[0]) and grundform.endswith(parts[1]) and not last_lemma.endswith(parts[1]):
                if all(not ((x in grundform[len(parts[0]):]) ^ (x in grundform)) for x in lemmata):
                    is_ableitung = True
                    break           
    # Prüfen auf Suffixe
    if not is_ableitung:
        for suffix in suffixes:
            if grundform.endswith(suffix) and not last_lemma.endswith(suffix):
                is_ableitung = True
                break
    if dim_checker(grundform,lemmata)[0]:
        return False,last_lemma
    
    return is_ableitung,last_lemma


def find_komposita(data:pd.DataFrame)->dict:
    """
    Funktionsweise:
        Erhalte einen pandas Dataframe, welcher mindestens die Spalten Grammatik, Grundform und Lemma enthalten muss. Finde alle Komposita der Grundformen des Dataframe und ordne diese in einem dict dem letzten Lemma zu.
        Dieses dict wird zurückgegeben
    Eingabe:
        data:pd.DataFrame = dies ist ein pandas Dataframe, welcher die Daten enthalten soll. Diese Daten müssen mindestens Grundform, Lemma und Grammatik beinhalten.
    Ausgabe:
        result:dict = dieses dictionary enthält alle Komposita des Dataframe data und ordnet diese dem letzten Lemma der Grundform zu.    
    """
    result=dict()
    data=data[data["Grammatik"].map(lambda x: str(x).startswith('S'))]  #Entferne alle Zeilen, die keine Substantive sind, Namen werden hier implizit entfernt
    data=data[['Grundform','Lemma']].drop_duplicates().groupby('Grundform').filter(lambda x: len(x)>1).groupby("Grundform")
    for grundform,group in data:
        lemmata=group['Lemma']
        lemma=find_last_lemma(grundform,lemmata)
        if " " not in grundform and "-" not in grundform: #entferne alle Wortgruppen
            if lemma in result:
                result[lemma].append(grundform)
            else:
                result[lemma]=[grundform]
    return result

    


def find_hapax_legomena(data:pd.DataFrame)->list[str]:
    """
    Funktionsweise:
        Findet alle Hapax Legomena eines Dataframe und gibt diese zurück
    Eingabe:
        data:pd.DataFrame= ein Dataframe, welcher mindestens Grundform und Lemma enthalten muss
    Ausgabe:
        result:list[str] = eine Liste mit allen Hapax Legomena des Dataframe
    """
    result=data[['Grundform','Lemma']].groupby(['Grundform','Lemma']).filter(lambda x: len(x) == 1)['Grundform'].drop_duplicates().tolist()
    return result


def find_diminuitive(data:pd.DataFrame)->dict:
    """
    Funktionsweise:
        Erhalte einen pandas Dataframe, welcher mindestens die Spalten Grammatik, Grundform und Lemma enthalten muss. Finde alle Diminuitiva der Grundformen des Dataframe und ordne diese in einem dict dem letzten Lemma zu.
        Dieses dict wird zurückgegeben
    Eingabe:
        data:pd.DataFrame = dies ist ein pandas Dataframe, welcher die Daten enthalten soll. Diese Daten müssen mindestens Grundform, Lemma und Grammatik beinhalten.
    Ausgabe:
        Diminuitiva:dict = dieses dictionary enthält alle Diminuitiva des Dataframe data und ordnet diese dem letzten Lemma der Grundform zu.
    """
    Diminuitiva=dict()
    data=data[data["Grammatik"].map(lambda x: str(x).startswith('S'))]  #Entferne alle Zeilen, die keine Substantive sind, Namen werden hier implizit entfernt
    data=data[['Grundform','Lemma']].drop_duplicates().groupby('Grundform')    #Dies generiert einen Dataframe, welcher zu jeder unterschiedlichen Grundform, alle Dazugehörigen Lemma anzeigt
    for grundform, group in data:
        lemmata=group['Lemma'].tolist()
        diminuitiv,lemma=dim_checker(grundform,lemmata)
        if diminuitiv:
            if lemma in Diminuitiva:
                Diminuitiva[lemma].append(grundform)
            else:
                Diminuitiva[lemma]=[grundform]
    return Diminuitiva

def find_last_lemma(grundform:str,lemmata:list)->str:
    """
    Funktionsweise:
        Finde das letzte Lemma eines Wortes, bei Angabe der Lemmata des Wortes.
    Eingabe:
        grundform:str = dies ist die grundform des Wortes
        lemmata:list = dies ist die Liste aller Lemmata des Wortes 
    Ausgabe:
        last_lemma:str = dies ist der String des gefundenen letzten Lemma, wenn das Wort keine Lemma hat, dann ist dieser String leer. 
    """
    grundform=grundform.lower()
 
    #check which lemma comes last in the word
    index=-1
    wort=grundform
    last_lemma=""
    if len(lemmata)==1:
        last_lemma=lemmata[0]
    else:
        for Lemma in lemmata:
            lemma=Lemma.lower()
            if lemma in wort:
                ind=wort.find(lemma)
                wort=wort.replace(lemma,"-" * len(lemma),1)
                if ind > index:
                    index=ind
                    last_lemma=Lemma
            else:
                temp=lemma
                while temp:
                    temp=temp[:-1]
                    if temp in wort:
                        ind=wort.find(temp)
                        wort=wort.replace(temp,"-" * len(temp),1)
                        if ind > index:
                            index=ind
                            last_lemma=Lemma
                        break
    return last_lemma

def dim_checker(grundform:str,lemmata:list)->tuple: 
    """
    Funktionsweise:
        hier wird überprüft ob ein Wort mit gegebenen Lemmata ein Diminuitiv ist und gibt das letzte Lemma des Wortes zurück.
    Eingabe:
        grundform:str= die Grundform des Wortes
        lemmata:list= Eine Liste mit allen Lemmata des Wortes
    Ausgabe:
        diminuitiv:boolean = Gibt an ob das Wort ein Diminuitiv ist
        last_lemma:str = das letzte Lemma des Wortes
    
    """
    grundform=grundform.lower()
 
    #check which lemma comes last in the word
    last_lemma=find_last_lemma(grundform,lemmata)
    #das letzte Lemma wurde gefunden


    diminuitiv=False
    if " " not in grundform and "-" not in grundform: #hier werden grundformen aus mehreren Worten oder Worten mit Bindestrichen entfernt
        for suffix in diminutive_suffixes:
            if grundform.endswith(suffix) and not last_lemma.endswith(suffix): #check ob das Suffix echt ist und nicht Teil des letzten Lemma
                #
                diminuitiv = True
                break
    
    return diminuitiv,last_lemma



def test1(name="Ableitungen.csv"):
    tmp_time=time.time()
    data= read_data()
    print("Finde Ableitungen")
    print("-----------------")
    ableitungen=find_ableitungen(data)
    print("Ableitungen gefunden")
    print("-----------------")
    print("Ableitungen werden in "+name+" gespeichert")
    print("-----------------")
    max_zeilen = max(len(liste) for liste in ableitungen.values())

    # Excel-Datei erstellen
    wb = Workbook()
    ws = wb.active
    ws.title = "Ableitungen"

    # Spaltenüberschriften schreiben
    headers = list(ableitungen.keys())
    ws.append(headers)

    # Zeilen schreiben, fehlende Werte mit None auffüllen
    for i in range(max_zeilen):
        zeile = [ableitungen[spalte][i] if i < len(ableitungen[spalte]) else None for spalte in headers]
        ws.append(zeile)

    # Datei speichern
    wb.save("Ableitungen.xlsx")
    with open(name,mode='w',encoding="utf-8" ,errors='replace' ) as file:
        writer=csv.writer(file)
        writer.writerow(["Lemma","Grundform"])
        for key,value in ableitungen.items():
            writer.writerow([key,value])
    print("Ableitungen gespeichert")
    print("-----------------")
    print("Ausführungsdauer Test 1: %s Sekunden" %(time.time()-tmp_time))
    print("-----------------")    

def test2(name="hapax_legomena.csv"):
    tmp_time=time.time()
    data=read_data()
    print("Finde Hapax Legomena")
    print("-----------------")
    hapax=find_hapax_legomena(data)
    print("Hapax Legomena gefunden")
    print("-----------------")
    print("Hapax Legomena werden in " +name+" gespeichert" )
    print("-----------------")

    with open(name, mode='w',encoding='utf-8',errors='replace') as file:
        writer=csv.writer(file)
        for value in hapax:
            writer.writerow([value])
    print("Hapax Legomena geschrieben")
    print("-----------------")
    print("Ausführungsdauer Test 2: %s Sekunden" %(time.time()-tmp_time))
    print("-----------------")    

def test3(name="Diminuitiva.csv"):
    tmp_time=time.time()
    data=read_data()
    print("Finde Diminuitiva")
    print("-----------------")
    diminuitiva=find_diminuitive(data)
    print("Diminuitiva gefunden")
    print("-----------------")
    print("Diminuitiva werden in " +name+" gespeichert" )
    print("-----------------")
    max_zeilen = max(len(liste) for liste in diminuitiva.values())
    # Excel-Datei erstellen
    wb = Workbook()
    ws = wb.active
    ws.title = "Diminuitiva"

    # Spaltenüberschriften schreiben
    headers = list(diminuitiva.keys())
    ws.append(headers)

    # Zeilen schreiben, fehlende Werte mit None auffüllen
    for i in range(max_zeilen):
        zeile = [diminuitiva[spalte][i] if i < len(diminuitiva[spalte]) else None for spalte in headers]
        ws.append(zeile)

    # Datei speichern
    wb.save("Diminuitiva.xlsx")
    with open(name, mode='w',encoding='utf-8',errors='replace') as file:
        writer=csv.writer(file)
        writer.writerow(["Lemma","Grundform"])
        for key,value in diminuitiva.items():
            writer.writerow([key,value])
    print("Diminuitiva geschrieben")
    print("-----------------")
    print("Ausführungsdauer Test 3: %s Sekunden" %(time.time()-tmp_time))
    print("-----------------")    


def test4(name="Komposita.csv"):
    tmp_time=time.time()
    data=read_data()
    print("Finde Komposita")
    print("-----------------")
    Komposita=find_komposita(data)
    print("Komposita gefunden")
    print("-----------------")
    print("Komposita werden in " +name+" gespeichert" )
    print("-----------------")
    max_zeilen = max(len(liste) for liste in Komposita.values())
    # Excel-Datei erstellen
    wb = Workbook()
    ws = wb.active
    ws.title = "Komposita"

    # Spaltenüberschriften schreiben
    headers = list(Komposita.keys())
    ws.append(headers)

    # Zeilen schreiben, fehlende Werte mit None auffüllen
    for i in range(max_zeilen):
        zeile = [Komposita[spalte][i] if i < len(Komposita[spalte]) else None for spalte in headers]
        ws.append(zeile)

    # Datei speichern
    wb.save("Komposita.xlsx")
    with open(name, mode='w',encoding='utf-8',errors='replace') as file:
        writer=csv.writer(file)
        writer.writerow(["Lemma","Grundform"])
        for key,value in Komposita.items():
            writer.writerow([key,value])
    print("Komposita geschrieben")
    print("-----------------")
    print("Ausführungsdauer Test 4: %s Sekunden" %(time.time()-tmp_time))
    print("-----------------")    


if __name__=="__main__":
    start_time = time.time()
    print("Execute")
    test1()
    test2()
    test3()
    test4()
    print("Done")
    print("Gesamtdauer: %s Sekunden" %(time.time()-start_time))