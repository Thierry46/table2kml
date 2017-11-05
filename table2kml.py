#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
"""
*********************************************************
Programme : table2kml.py
Github : https://github.com/Thierry46/table2kml
Auteur : Thierry Maillard (TMD)
Date : 27 - 5/11/2017

Role : Convertir les coordonées informations de localisation contenues
        dans un fichier de donnees (Excel97 ou CSV)
        au format KML importable dans Geoportail.

Prerequis :
- Python v3.xxx : a télécharger depuis : https://www.python.org/downloads/
- module simplekml : sudo python3 -m pip install simplekml
- module xlrd : sudo python3 -m pip install xlrd

Usage : table2kml.py [-h] [-v] [Chemin xls]
Sans paramètre, lance une IHM, sinon fonctionne en batch avec 1 parametre.

Environnement :
    Ce programme teste son environnement (modules python disponibles)
    et s'y adaptera.
    tkinter : pour IHM : facultatif (mode batch alors seul)
    xlrd : pour lire le fichier Excel (obligatoire pour traiter fichier .xls en entrée)
    simplekml : pour ecrire le fichier resultat kml (obligatoire)

Parametres :
    -h ou --help : affiche cette aide.
    -v ou --verbose : mode bavard
    Nom d'un fichier de données Excel .xls ou .csv (mode batch)
    Titre du calque codé dans le fichier KML : Ex.: "Dolmen Adrien" (mode batch)
    URL du picto : (mode batch)
    Ex.: https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png

Sortie :
    - Fichier de même nom que fichier d'entrée mais avec extension .kml

Exemples de lancement par ligne de commande sous Mac :
=====================
Conversion batch d'un fichier CSV
cd dossier application
./table2kml.py Dolmen_csv_v0.6.csv "Dolmens Adrien" https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png
Conversion batch d'un fichier Excel 97 :
./table2kml.py Dolmen_v0.6.xls "Dolmens Adrien" https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png
Lancement IHM :
./table2kml.py

Sous Windows :
Lancement IHM : double-cliquer sur table2kml.py
lancement Batch :
Ouvrir une fenêtre de commande : Windows + E, commande cmd
cd dossier appli
commandes identiques à celles au dessus sans ./

Modifications : voir https://github.com/Thierry46/table2kml

Copyright 2017 Thierry Maillard
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
Contact me at thierry.maillard500n@orange.fr
*********************************************************
"""
import sys
import getopt
import math
import time
import imp
import os
import os.path
import platform
import re
import time
import getpass

##################################################
# main function
##################################################
def main(argv=None):
    VERSION = 'v1.3 - 5/11/2017'
    NOM_PROG = 'table2kml.py'
    isVerbose = False
    title = NOM_PROG + ' - ' + VERSION + " sur " + platform.system() + " " + platform.release()
    print(title)

    # Test environnement
    # Test presence des modules tkinter, numpy et matplotlib
    canUseGUI = True
    canUseXLS = True
    try :
        imp.find_module('tkinter')
        import tkinter
        print("Module tkinter OK : GUI allowed")
    except ImportError as exc:
        print("Warning ! module tkinter pas disponible -> utiliser mode batch :", exc)
        canUseGUI = False
    try :
        imp.find_module('xlrd')
        print("Module xlrd OK : .xls files supported")
    except ImportError as exc:
        print("Warning ! module xlrd pas disponible -> fichiers .xls non autorisé :", exc)
        canUseXLS = False
    try :
        imp.find_module('simplekml')
        print("Module simplekml OK : export en KML autorisé")
    except ImportError as exc:
        print(__doc__)
        print("Erreur ! module simplekml pas disponible :", exc)
        sys.exit(1)

    if argv is None:
        argv = sys.argv

    # parse command line options
    dirProject = os.path.dirname(os.path.abspath(sys.argv[0]))
    try:
        opts, args = getopt.getopt(argv[1:], "hv", ["help", "verbose"])
    except getopt.error as msg:
        print(msg)
        print("To get help use --help ou -h")
        sys.exit(1)
    # process options
    for o, a in opts:
        if o in ("-h", "--help"):
            print(__doc__)
            sys.exit(0)

        if o in ("-v", "--verbose"):
            isVerbose = True
            print("Mode verbose : bavard pour debug")

    if len(args) < 1:
        if canUseGUI:
            import tkinter
            print("Lancement de l'IHM...")
            root = tkinter.Tk()
            table2kmlGUI(root, dirProject, title, canUseXLS, isVerbose)
            root.mainloop()
        else:
            print(__doc__)
            print("IHM non disponible : tkinter pas installe !")
            print("Utilisez le mode batch et passez le fichier à traiter au programme")
            sys.exit(2)

    else: # Batch
        if len(args) == 3:
            processFile(canUseXLS, args[0], args[1], args[2], isVerbose)
        else:
            print(__doc__)
            print("Nombre de paramètre invalide : 3 nécessaires : fichier titre picto")
            sys.exit(1)

    print('End table2kml.py', VERSION)
    sys.exit(0)

def processFile(canUseXLS, pathFicTable, titleKML, URLPicto, isVerbose):
    listInfoRead = []
    pathKMLFile = ""
    """ Convertit un fichier passé en paramètre en un fichier KML """
    if canUseXLS and pathFicTable.endswith(".xls"):
        listMessage, listInfoRead = readExcel(pathFicTable, isVerbose)
        pathKMLFile = pathFicTable.replace(".xls", ".kml")
    elif pathFicTable.endswith(".csv"):
        listMessage, listInfoRead = readCSV(pathFicTable, isVerbose)
        pathKMLFile = pathFicTable.replace(".csv", ".kml")
    else:
        raise ValueError("Extension du fichier non supporté :" +
                          os.path.basename(pathFicTable) +
                          " extension supportées : .xls")

    genKMLFiles(listInfoRead, titleKML, URLPicto, pathKMLFile, isVerbose)
    return listMessage, listInfoRead

def readExcel(pathFicTable, isVerbose):
    """ Recupère les infos de localisation dans le fichier Excel"""
    import xlrd
    EXT_FIC_OK = ".xls"
    LIGNE_START = 2
    dictData = {
        'Nom' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Lat' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Lon' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Lieu' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Commune' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Tumulus' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Orthostats' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Table' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Classement' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Remarques' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Pages Web' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'OSM' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None}
            }

    if len(pathFicTable) == 0:
        raise ValueError("Nom fichier vide !" )

    if not pathFicTable.endswith(EXT_FIC_OK):
        raise ValueError("Nom fichier incorrect : " +
                          os.path.basename(pathFicTable) +
                          " : devrait finir par l'extension .xls")

    print("Lecture de", pathFicTable, "...")
    # ouverture du fichier Excel 
    workbook = xlrd.open_workbook(filename= pathFicTable, on_demand=True)
    if isVerbose:
        print("Liste des feuilles du classeur :", workbook.sheet_names())

    # ouverture des données dans la 1ere feuille
    sheetData = workbook.sheet_by_index(0)

    # Recherche des noms de colonnes à extraire
    titleRow = [title.strip() for title in sheetData.row_values(0)]
    if isVerbose:
        print("Titres des colonnes :", titleRow)
    for numCol, title in enumerate(titleRow):
        for startColumn in dictData:
            if title.startswith(startColumn):
                dictData[startColumn]['numCol'] = numCol
                dictData[startColumn]['nomCol'] = title

    # test présence colonnes obligatoires et lecture des colonnes présentes
    for startColumn in dictData:
        if dictData[startColumn]['obligatoire'] and dictData[startColumn]['numCol'] == -1:
            raise ValueError("Aucune colonne commençant par " + startColumn + " trouvée !")
        if isVerbose:
            if dictData[startColumn]['numCol'] != -1:
                print("Num colonne Nom :", dictData[startColumn]['numCol'], ": ",
                      dictData[startColumn]['nomCol'])
        if dictData[startColumn]['numCol'] != -1:
            dictData[startColumn]['dataCol'] = [
                str(value).strip() for value in sheetData.col_values(dictData[startColumn]['numCol'])]

    #Extraction des liens WEB
    for numLigne, urlString in enumerate(dictData['Pages Web']['dataCol']):
        if "http" not in urlString:
            link = sheetData.hyperlink_map.get((numLigne, dictData['Pages Web']['numCol']))
            url = '' if link is None else link.url_or_path
            dictData['Pages Web']['dataCol'][numLigne] = url
 
    # Formatage et contrôle des donnees utiles
    listInfoRead = []
    listMessage = []
    for numLigne, nomDolmen in enumerate(dictData['Nom']['dataCol']):
        ligneOK = True
        messageInfos = {'numLigne':numLigne+1}

        if numLigne < LIGNE_START - 1:
            ligneOK = False
            messageInfos['texte'] = "ignorée car ligne de titre"

        if ligneOK and len(nomDolmen) == 0 :
            ligneOK = False
            messageInfos['texte'] = "ignorée car champ " + dictData['Nom']['nomCol'] + " vide"

        # Verif et conversion champs Longitude et Latitude
        coordValue = {}
        for field in ('Lon', 'Lat'):
            if ligneOK and (numLigne >= len(dictData[field]['dataCol']) or
                             len(dictData[field]['dataCol'][numLigne]) == 0) :
                  ligneOK = False
                  messageInfos['texte'] = "ignorée car champ " + dictData[field]['nomCol'] + " vide"
            if ligneOK:
                  try:
                      coordValue[field] = convertCoord(dictData[field]['dataCol'][numLigne])
                  except ValueError:
                      ligneOK = False
                      messageInfos['texte'] = "ignorée car champ " + \
                                        dictData[field]['nomCol'][numLigne] + \
                                        " incorrect : " + dictData[field]['dataCol'][numLigne]

        # Construction du champ description
        description = "<h1>Informations</h1>" + '\n'
        for field in ('Commune', 'Lieu', 'Lat', 'Lon', 'Classement', 'Remarques',
                      'Tumulus', 'Orthostats', 'Table', 'OSM', 'Pages Web'):
            if ligneOK and dictData[field]['numCol'] != -1 and \
                numLigne < len(dictData[field]['dataCol']) and \
                len(dictData[field]['dataCol'][numLigne]) != 0 :
                description += "<b>" + dictData[field]['nomCol'] + "</b> : "

                # Champs particuliers
                if field == 'Commune':
                    nomCommune = dictData[field]['dataCol'][numLigne]
                    url = 'https://fr.wikipedia.org/wiki/' + nomCommune
                    description += '<a href="' + url + '" target="_blank">' + nomCommune + \
                                   '</href><br/>\n'

                elif field == 'Lat' or field == 'Lon':
                    # Ecrit dans le champ description les coordonnées converties
                    description += str(coordValue[field]) + '<br/>\n'

                elif field == 'Pages Web':
                    url = dictData[field]['dataCol'][numLigne]
                    description += '<a href="' + url + '" target="_blank">Infos WEB</href><br/>\n'

                else:
                    description += dictData[field]['dataCol'][numLigne]
                    description += '<br/>\n'

        # Enregistrement des valeurs utiles dans la structure résultat
        if ligneOK:
            listInfoRead.append({'numLigne':numLigne+1, 'nom':nomDolmen,
                              'Commune' : dictData['Commune']['dataCol'][numLigne],
                              'latitude':coordValue['Lat'],
                              'longitude':coordValue['Lon'],
                              'description':description
                              })
        else:
            listMessage.append(messageInfos)

    print(len(listInfoRead), "dolmens enregistrés,", len(listMessage), "lignes ignorées.")

    if isVerbose:
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])

    return listMessage, listInfoRead

def readCSV(pathFicTable, isVerbose):
    """ Recupère les infos de localisation dans le fichier Excel"""
    import csv
    EXT_FIC_OK = ".csv"
    LIGNE_START = 2
    dictData = {
        'Nom' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Lat' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Lon' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Lieu' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Commune' : {'obligatoire':True, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Tumulus' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Orthostats' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Table' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Classement' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Remarques' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Pages Web' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'OSM' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None}
        }

    listInfoRead = []
    listMessage = []

    if len(pathFicTable) == 0:
        raise ValueError("Nom fichier vide !")

    if not pathFicTable.endswith(EXT_FIC_OK):
        raise ValueError("Nom fichier incorrect : " +
                         os.path.basename(pathFicTable) +
                         " : devrait finir par l'extension .xls")

    print("Lecture de", pathFicTable, "...")
    # Analyse du fichier CSV
    with open(pathFicTable, newline='') as csvfile:
        sample = csvfile.read(1024)
        sniffer = csv.Sniffer()
        if not sniffer.has_header(sample):
            raise ValueError("Impossible de trouver une entête dans le fichier !")
        elif isVerbose:
            print("Ligne d'entête détectée.")

        dialect = sniffer.sniff(sample)
        if dialect is  None:
            raise ValueError("Impossible de trouver le dialecte CSV du fichier !")
        elif isVerbose:
            print("Dialect CSV détecté")
        csvfile.seek(0)

        # Lecture du fichier dans dictionnaire
        reader = csv.DictReader(csvfile, dialect=dialect)

        # Analyse ligne de titre
        titleRow = reader.fieldnames
        if isVerbose:
            print("Titres des colonnes :", titleRow)
        for numCol, title in enumerate(titleRow):
            for startColumn in dictData:
                if title.startswith(startColumn):
                    dictData[startColumn]['numCol'] = numCol
                    dictData[startColumn]['nomCol'] = title

        # test présence colonnes obligatoires et lecture des colonnes présentes
        for startColumn in dictData:
            if dictData[startColumn]['obligatoire'] and dictData[startColumn]['numCol'] == -1:
                raise ValueError("Aucune colonne commençant par " + startColumn + " trouvée !")
            if isVerbose and dictData[startColumn]['numCol'] != -1:
                print("Num colonne Nom :", dictData[startColumn]['numCol'], ": ",
                      dictData[startColumn]['nomCol'].strip())

        # Analyse et enregistrement valeurs
        for row in reader:
            ligneOK = True
            messageInfos = {'numLigne':reader.line_num}

            if ligneOK and len(row[dictData['Nom']['nomCol']]) == 0:
                ligneOK = False
                messageInfos['texte'] = "ignorée car champ " + row['Nom'] + " vide"

            # Verif et conversion champs Longitude et Latitude
            coordValue = {}
            for field in ('Lon', 'Lat'):
                if len(row[dictData[field]['nomCol']]) == 0 :
                    ligneOK = False
                    messageInfos['texte'] = "ignorée car champ " + dictData[field]['nomCol'] + \
                                            " vide"
                if ligneOK:
                    try:
                       coordValue[field] = convertCoord(row[dictData[field]['nomCol']])
                    except ValueError:
                        ligneOK = False
                        messageInfos['texte'] = "ignorée car champ " + \
                                    dictData[field]['nomCol'] + " incorrect : " + \
                                    row[dictData[field]['nomCol']]

            # Construction du champ description
            description = "<h1>Informations</h1>" + '\n'
            for field in ('Commune', 'Lieu', 'Lat', 'Lon', 'Classement', 'Remarques',
                           'Tumulus', 'Orthostats', 'Table', 'OSM', 'Pages Web'):
                if ligneOK and dictData[field]['numCol'] != -1 and \
                    len(row[dictData[field]['nomCol']]) != 0 :
                    description += "<b>" + dictData[field]['nomCol'].strip() + "</b> : "
                    # Champs particuliers
                    if field == 'Commune':
                        nomCommune = row[dictData[field]['nomCol']]
                        url = 'https://fr.wikipedia.org/wiki/' + nomCommune
                        description += '<a href="' + url + '" target="_blank">' + nomCommune + \
                                        '</href><br/>\n'

                    elif field == 'Lat' or field == 'Lon':
                        # Ecrit dans le champ description les coordonnées converties
                        description += str(coordValue[field]) + '<br/>\n'

                    elif field == 'Pages Web':
                        url = row[dictData[field]['nomCol']]
                        description += '<a href="' + url + \
                                        '" target="_blank">Infos WEB</href><br/>\n'

                    else:
                        description += row[dictData[field]['nomCol']]
                        description += '<br/>\n'

            # Enregistrement des valeurs utiles dans la structure résultat
            if ligneOK:
                listInfoRead.append({'numLigne':reader.line_num,
                                    'nom':row[dictData['Nom']['nomCol']],
                                    'Commune' : row[dictData['Commune']['nomCol']],
                                    'latitude':coordValue['Lat'],
                                    'longitude':coordValue['Lon'],
                                    'description':description
                                    })
            else:
                listMessage.append(messageInfos)

    print(len(listInfoRead), "dolmens enregistrés,", len(listMessage), "lignes ignorées.")

    if isVerbose:
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])

    return listMessage, listInfoRead

def convertCoord(coord):
    """ Convertit une chaine de coordonnées d'angle en un réel
        La chaine d'entrée peut contenir une valeur sexagésimale """
    valFloat = -1.0
    try:
        valFloat = float(coord)
    except ValueError:
        # re OK pour expression du type 44°51'37" ou 1°51'37" ou 1°51'37"" ou 1°51'37" "
        m = re.search(r'(?P<degres>[\d]{1,2})°(?P<minutes>[\d]{2})\'(?P<secondes>[\d]{2})"',
                      coord)
        if m:
            valFloat = float(m.group('degres')) + float(m.group('minutes')) / 60. + \
                       float(m.group('secondes')) / 3600.
        else:
            raise
    return valFloat

def genKMLFiles(listInfoRead, titleKML, URLPicto, pathKMLFile, isVerbose):
    """ genere les fichiers de sortie KML"""

    import simplekml
    # Ref simplekml:http://www.simplekml.com/en/latest/reference.html

    if not URLPicto.startswith("http"):
        raise ValueError("URL du picto incorrecte :" +
                         URLPicto +
                         " : devrait commencé par http")

    print("Ecriture des résultats dans", pathKMLFile, "...")
    titleKML = titleKML + " " + time.strftime("%d/%m/%y")
    kml = simplekml.Kml(name=titleKML)

    # Style icone et couleur du texte pour tous les dolmens
    # Ref couleur:http://www.simplekml.com/en/latest/constants.html#color
    styleIconDolmen = simplekml.Style()
    styleIconDolmen.labelstyle.color = simplekml.Color.cadetblue
    styleIconDolmen.iconstyle.icon.href = URLPicto

    for dolmen in listInfoRead:
        point = kml.newpoint(name=dolmen['nom'],
                             description='<![CDATA[' + dolmen['description'] + ']]>\n',
                             coords=[(str(dolmen['longitude']), str(dolmen['latitude']))])
        point.style = styleIconDolmen

    kml.save(pathKMLFile)
    print(str(len(listInfoRead)), "dolmens écrits dans", pathKMLFile)

############
class table2kmlGUI():
    """
    A GUI for table2kml script.
    """
    def __init__(self, root, dirProject, title, canUseXLS, isVerbose):
        """
        Constructor
        Define all GUIs widgets
        parameters :
            - root : main window
            - title : Windows title
            - isVerbose : If true, environment is OK for plotting
        """
        import tkinter

        self.URL_DOLMEN = "https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png"

        self.root = root
        mainFrame = tkinter.Frame(self.root)
        self.dirProject = dirProject
        self.canUseXLS = canUseXLS
        self.isVerbose = isVerbose
        self.listMessage = []
        self.listInfoRead = []

        self.root.title(title)

        # Frame des paramètres de configuration
        inputFrame = tkinter.LabelFrame(mainFrame, text="Parametres")
        readConfigButton = tkinter.Button(inputFrame,
                                               text="Fichier ...",
                                               command=self.fileChooserCallback,
                                               bg='green')
        readConfigButton.grid(row=0, column=0)
        self.pathInputFileVar = tkinter.StringVar()
        pathInputFileEntry = tkinter.Entry(inputFrame,
                                          textvariable=self.pathInputFileVar,
                                          width=60)
        pathInputFileEntry.grid(row=0, column=1, sticky=tkinter.W)

        titleKMLButton = tkinter.Button(inputFrame, text="Titre KML :",
                                  command=self.initTitleKMLCallback)
        titleKMLButton.grid(row=1, column=0)
        self.titleKMLVar = tkinter.StringVar()
        self.titleKMLEntry = tkinter.Entry(inputFrame,
                                           textvariable=self.titleKMLVar,
                                           width=60)
        self.titleKMLEntry.grid(row=1, column=1)

        titlePictoButton = tkinter.Button(inputFrame, text="URL picto :",
                                          command=self.initPictoCallback)
        titlePictoButton.grid(row=2, column=0)
        self.urlPictoVar = tkinter.StringVar()
        self.urlPictoEntry = tkinter.Entry(inputFrame,
                                           textvariable=self.urlPictoVar,
                                           width=60)
        self.urlPictoEntry.grid(row=2, column=1)

        tkinter.Button(inputFrame, text="Traiter le fichier",
                       command=self.launchInputFileReader,
                       fg='red').grid(row=3, column=0, columnspan=2)
        inputFrame.pack(side = tkinter.TOP, fill="both", expand="yes")

        # Pour affichage des messages de lecture
        messageFrame = tkinter.LabelFrame(mainFrame, text="Affichage des messages de lecture")
        self.messagesListbox = tkinter.Listbox(messageFrame,
                                               background="green yellow",
                                               height=10, width=70)
        self.messagesListbox.grid(row=0, columnspan=2)
        scrollbarRightMessages = tkinter.Scrollbar(messageFrame, orient=tkinter.VERTICAL,
                                   command=self.messagesListbox.yview)
        scrollbarRightMessages.grid(row=0, column=2, sticky=tkinter.W+tkinter.N+tkinter.S)
        self.messagesListbox.config(yscrollcommand=scrollbarRightMessages.set)
        scrollbarBottomMessages = tkinter.Scrollbar(messageFrame, orient=tkinter.HORIZONTAL,
                                        command=self.messagesListbox.xview)
        scrollbarBottomMessages.grid(row=1, columnspan=2, sticky=tkinter.N+tkinter.E+tkinter.W)
        self.messagesListbox.config(xscrollcommand=scrollbarBottomMessages.set)
        messageFrame.pack(side = tkinter.TOP, fill="both", expand="yes")

        # Pour affichage des dolmens lus
        dolmensFrame = tkinter.LabelFrame(mainFrame, text="Dolmens trouvés dans le fichier")
        self.dolmensListbox = tkinter.Listbox(dolmensFrame,
                                       background="light blue",
                                       height=10, width=70)
        self.dolmensListbox.grid(row=0, columnspan=2)
        scrollbarRightDolmens = tkinter.Scrollbar(dolmensFrame, orient=tkinter.VERTICAL,
                                                  command=self.dolmensListbox.yview)
        scrollbarRightDolmens.grid(row=0, column=2, sticky=tkinter.W+tkinter.N+tkinter.S)
        self.dolmensListbox.config(yscrollcommand=scrollbarRightDolmens.set)
        scrollbarBottomDolmens = tkinter.Scrollbar(dolmensFrame, orient=tkinter.HORIZONTAL,
                                                   command=self.dolmensListbox.xview)
        scrollbarBottomDolmens.grid(row=1, columnspan=2, sticky=tkinter.N+tkinter.E+tkinter.W)
        self.dolmensListbox.config(xscrollcommand=scrollbarBottomDolmens.set)
        dolmensFrame.pack(side = tkinter.TOP, fill="both", expand="yes")

        statusFrame = tkinter.LabelFrame(mainFrame, text="Messages")
        self.messageLabel = tkinter.Label(statusFrame, text="Pret !")
        self.messageLabel.pack(side = tkinter.LEFT)
        statusFrame.pack(side = tkinter.BOTTOM, fill="both", expand="yes")
        mainFrame.pack()

    ################
    # Callbacks
    ################
        
    def fileChooserCallback(self) :
        """
        Callback for button that launch file chooser for input file.
        """
        from tkinter import filedialog

        # Le dernier répertoire du fichier de travail est relu pour
        #        faciliter les traitements répétitifs.
        lastDir = "."
        try :
            with open(os.path.join(self.dirProject, "table2kmlDir.txt"), 'r') as hDirFile:
                lastDir = hDirFile.readline().strip()
                hDirFile.close()      
        except IOError:
            pass
        fileName = filedialog.askopenfilename(
                        parent=self.root,
                        initialdir=lastDir,
                        filetypes = [("Fichier CSV","*.csv"), ("Fichier Excel","*.xls"), ("All", "*")],
                        title="Selectionnez un fichier de configuration")
        if fileName != "":
            # Le dernier répertoire du fichier de travail est stocké pour
            #        faciliter les traitements répétitifs.
            try :
                with open(os.path.join(self.dirProject, "table2kmlDir.txt"), 'w') as hDirFile:
                    hDirFile.write(os.path.dirname(fileName) + "\n")
                    hDirFile.close()      
            except IOError:
                self.setMessageLabel(str(exc), True)
            self.pathInputFileVar.set(fileName)

    def initTitleKMLCallback(self) :
        """
        Callback for setting default value for KML title.
        """
        self.titleKMLVar.set("Calque de " + getpass.getuser())

    def initPictoCallback(self) :
        """
        Callback for setting default value for KML title.
        """
        self.urlPictoVar.set(self.URL_DOLMEN)

    def launchInputFileReader(self, event=None):
        """ Update list according input files
            v2.0:Thread use"""
        self.setMessageLabel("Lecture du fichier en cours, veuillez patienter...")
        import tkinter

        pathFicTable = self.pathInputFileVar.get()
        try :
            self.listMessage, self.listInfoRead = \
                processFile(self.canUseXLS, pathFicTable,
                            self.titleKMLEntry.get(), self.urlPictoEntry.get(),
                            self.isVerbose)
            if len(self.listInfoRead) == 0:
                raise ValueError("Aucun dolmen trouvé dans le fichier !")

            # Update message list
            self.messagesListbox.delete(0, tkinter.END)
            for message in self.listMessage:
                self.messagesListbox.insert(tkinter.END,
                                            "Ligne " + str(message['numLigne']) +
                                            " : " + message['texte'])
            # Update dolmen list
            self.dolmensListbox.delete(0, tkinter.END)
            for dolmen in self.listInfoRead:
                self.dolmensListbox.insert(tkinter.END,
                                           dolmen['Commune'] + " : " + dolmen['nom'])

            self.setMessageLabel("Fichier converti : " + str(len(self.listMessage)) +
                                        " messages, " + str(len(self.listInfoRead)) +
                                        " dolmens dans : " +
                                        os.path.basename(pathFicTable).replace('.xls', '.kml'))
        except IOError as exc:
            self.setMessageLabel(str(exc), isError=True)
        except OSError as exc:
            self.setMessageLabel(str(exc), isError=True)
        except ValueError as exc:
            self.setMessageLabel(str(exc), isError=True)

    def setMessageLabel(self, message, isError=False) :
        """ set a message for user. """
        if isError:
            message = "Probleme : " + message
        self.messageLabel['text'] =  message
        if isError:
            print(message)

#to be called as a script:python table2kml.py or table2kml.py
if __name__ == "__main__":
    main()
