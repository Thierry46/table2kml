#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
"""
*********************************************************
Programme : table2kml.py
Github : https://github.com/Thierry46/table2kml
Auteur : Thierry Maillard (TMD)
Date : 27/11/2017 - 7/12/2017

Role : Convertir les coordonnées et informations contenues dans un fichier
        au format KML importable dans Geoportail.
        Deux formats pour le fichier d'entrée sont supportés :
        - Excel 97 (1ère feuille du classeur)
        - format CSV de format plus souple à utiliser de préférence.

        Le fichier d'entrée doit contenir des colonnes commençant par :
        - Nom : nom à afficher dans le picto
        - Lat : latitude en degré sexagésimal ou décimal
        - Lon : longitude en degré sexagésimal ou décimal
        Les autres colonnes seront affichées dans la bulle d'info.
        Les colonnes dont les titres commencent par - seront ignorées.

Prerequis :
- Python v3.xxx : a télécharger depuis : https://www.python.org/downloads/
- module simplekml : sudo python3 -m pip install simplekml
- module xlrd (pour fichier Excel) : sudo python3 -m pip install xlrd

Environnement :
Ce programme teste son environnement (modules python disponibles)
et s'y adaptera.
tkinter : pour IHM : facultatif (mode batch alors seul)
xlrd : pour lire le fichier Excel (obligatoire pour traiter fichier .xls en entrée)
simplekml : pour ecrire le fichier resultat kml (obligatoire)

Usage : table2kml.py [-h] [-v] [Chemin_fichier Nom_calque [url_picto]]
Sans paramètre, lance une IHM, sinon fonctionne en batch avec 1 parametre.
Parametres :
    -h ou --help : affiche cette aide.
    -v ou --verbose : mode bavard
    Nom d'un fichier de données Excel .xls ou .csv (mode batch)
    Titre du calque codé dans le fichier KML : Ex.: "Dolmen Adrien" (mode batch)
    URL du picto : (facultatif en mode batch)
    Ex.:
    https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png (Base Adrien)
    https://upload.wikimedia.org/wikipedia/commons/2/2b/1_dfs.png (carte Wkp)
    https://upload.wikimedia.org/wikipedia/commons/c/c5/3_slmb.png (Article Wkp)
    https://upload.wikimedia.org/wikipedia/commons/c/cc/5_ldf.png (T4T35)

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
    VERSION = 'v3.1 - 14/12/2017'
    NOM_PROG = 'table2kml.py'
    isVerbose = False
    title = NOM_PROG + ' - ' + VERSION + " sur " + platform.system() + " " + platform.release() + \
        " - Python : " + platform.python_version()
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
        if len(args) >=2 and len(args) <= 3:
            URLPicto = ""
            if len(args) == 3:
                URLPicto = args[2]
            processFile(canUseXLS, args[0], args[1], URLPicto, isVerbose)
        else:
            print(__doc__)
            print("Nombre de paramètre invalide : 2 nécessaires et 1 facultatif :")
            print("fichier titre [URLpicto]")
            sys.exit(1)

    print('End table2kml.py', VERSION)
    sys.exit(0)

def processFile(canUseXLS, pathFicTable, titleKML, URLPicto, isVerbose):
    """ Convertit un fichier passé en paramètre en un fichier KML """
    listInfoRead = []
    titleRow = []
    rowData = []
    neededColumns = ['Nom', 'Lat', 'Lon']
    pathKMLFile = ""
    if canUseXLS and pathFicTable.endswith(".xls"):
        titleRow, rowData = readExcel(pathFicTable, isVerbose)
        pathKMLFile = pathFicTable.replace(".xls", ".kml")
    elif pathFicTable.endswith(".csv"):
        titleRow, rowData = readCSV(pathFicTable, isVerbose)
        pathKMLFile = pathFicTable.replace(".csv", ".kml")
    else:
        raise ValueError("Extension du fichier non supporté :" +
                          os.path.basename(pathFicTable) +
                          " extension supportées : .xls")

    listMessage, listInfoRead = formatData(titleRow, rowData, neededColumns, isVerbose)
    genKMLFiles(listInfoRead, titleKML, URLPicto, pathKMLFile, isVerbose)
    return listMessage, listInfoRead

def readExcel(pathFicTable, isVerbose):
    """ Recupère les infos de localisation dans le fichier Excel"""
    import xlrd
    EXT_FIC_OK = ".xls"
    neededColumns = ['Nom', 'Lat', 'Lon']
    titleRow = []
    rowData = []

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
    if isVerbose:
        print("Analyse de la feuille 0  :", sheetData.name)

    # Enregistrement contenu de la table dans la structure rowData
    titleRow = sheetData.row_values(0)
    rowData = []
    for numRow in range(1, sheetData.nrows):
        rowCols = {}
        for numCol in range(sheetData.ncols):
            rowCols[sheetData.cell_value(0, numCol)] = sheetData.cell_value(numRow, numCol)
            # Extraction des liens WEB
            link = sheetData.hyperlink_map.get((numRow, numCol))
            if link is not None:
                rowCols[sheetData.cell_value(0, numCol)] = link.url_or_path
        rowData.append(rowCols)

    return titleRow, rowData

def readCSV(pathFicTable, isVerbose):
    """ Recupère les infos de localisation dans le fichier Excel"""
    import csv
    EXT_FIC_OK = ".csv"
    titleRow = []
    rowData = []

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

        try:
            if not sniffer.has_header(sample):
                raise ValueError("Impossible de trouver une entête dans le fichier !")
            elif isVerbose:
                print("Ligne d'entête détectée.")
            dialect = sniffer.sniff(sample)
            if dialect is  None:
                raise ValueError("Impossible de trouver le dialecte CSV du fichier !")
            elif isVerbose:
                print("Dialect CSV détecté")
        except csv.Error as exc:
            print("Erreur : ", str(exc))
            print("Essai dialect Excel avec séparateur de champ ;")
            # Enregistre ce dialecte auprès du module csv
            csv.register_dialect('excel-fr', delimiter=';')
            dialect = 'excel-fr'

        # Lecture du fichier dans dictionnaire
        csvfile.seek(0)
        reader = csv.DictReader(csvfile, dialect=dialect)

        # Enregistrement contenu de la table dans la structure rowData
        titleRow = reader.fieldnames
        for row in reader:
            rowData.append(row)
    return titleRow, rowData

def formatData(titleRow, rowData, neededColumns, isVerbose):
    """ Formatage et contrôle des donnees utiles """
    listInfoRead = []
    listMessage = []
    regexpSite = re.compile(r'^http[s]?://(?P<siteName>.+?)/.*?(?P<id>[\w=. ]+)$')

    # Détermination colonnes utiles
    titleRowUsed = checkNeededColumns(titleRow, neededColumns, isVerbose)

    for numLigne, row in enumerate(rowData):
        ligneOK = True
        messageInfos = {'numLigne':numLigne+1}

        # Check neededColumns[0]
        fieldName, nomElement = getFirstFieldStartingBy(row, neededColumns[0])
        if ligneOK and len(nomElement) == 0 :
            ligneOK = False
            messageInfos['texte'] = "ignorée car champ " + fieldName + " vide"

        # Verif et conversion champs Longitude et Latitude
        coordValue = {}
        for field in (neededColumns[1], neededColumns[2]):
            fieldName, value = getFirstFieldStartingBy(row, field)
            if ligneOK and len(value) == 0 :
                  ligneOK = False
                  messageInfos['texte'] = "ignorée car champ " + fieldName + " vide"
            if ligneOK:
                  try:
                      coordValue[fieldName] = convertCoord(value)
                  except ValueError:
                      ligneOK = False
                      messageInfos['texte'] = "ignorée car champ " + fieldName + \
                                              " incorrect : " + value

        # Construction du champ description
        description = "<h1>Informations</h1>" + '\n'
        fieldCommune = ""
        for field in titleRowUsed:
            # Place name is not written in description info balloon
            if ligneOK and not field.startswith(neededColumns[0]):
                value = str(row[field]).strip()
                if len(value) != 0 :
                    description += "<b>" + field.strip() + "</b> : "

                    # Champs particuliers
                    if field.startswith('Commune'):
                        fieldCommune = value
                        value = 'https://fr.wikipedia.org/wiki/' + value
                    elif field == getFirstFieldStartingBy(row, neededColumns[1])[0] or \
                        field == getFirstFieldStartingBy(row, neededColumns[2])[0]:
                        # Ecrit dans le champ description les coordonnées converties
                        value = str(coordValue[field])

                    if value.startswith("http"):
                        value = formateURL(value, regexpSite)
                    description += value + '<br/>\n'

        # Enregistrement des valeurs utiles dans la structure résultat
        if ligneOK:
            listInfoRead.append({'numLigne':numLigne+1,
                                'nom':nomElement.strip(),
                                'Commune':fieldCommune,
                                'latitude':coordValue[getFirstFieldStartingBy(row, neededColumns[1])[0]],
                                'longitude':coordValue[getFirstFieldStartingBy(row, neededColumns[2])[0]],
                                'description':description
                                })
        else:
            listMessage.append(messageInfos)

    if isVerbose:
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])
    print(len(listInfoRead), "éléments enregistrés,", len(listMessage), "lignes ignorées.")

    return listMessage, listInfoRead

def  checkNeededColumns(allColumnNames, neededColumns, isVerbose):
    """ Verif présence colonnes obligatoires dans titres
        Suppression colonne commençant par -
        Retourne la liste des titres et les noms complets des colonnes obligatoires """

    # Elimination des colonnes commençant par -
    titleRow = [title for title in allColumnNames if not title.startswith('-')]
    if isVerbose:
        print("Titres des colonnes :", titleRow)

    # Test nombre de colonnes
    if len(titleRow) < len(neededColumns):
        raise ValueError("Au moins " + len(neededColumns) + " nécessaires !")

    for startColumn in neededColumns:
        colFound = False
        for colName in titleRow:
            if colName.startswith(startColumn):
                colFound = True
                break
        if not colFound:
            raise ValueError("Aucune colonne commençant par " + startColumn + " trouvée !")

    if isVerbose:
        print("Titres des colonnes obligatoires OK :", neededColumns)
    return titleRow

def getFirstFieldStartingBy(row, startName):
    """ in the dictionary row, get first field name starting with startName
        and its value """
    value = "?"
    for fieldName in row.keys():
        if fieldName.startswith(startName):
            value = row[fieldName]
            break
    if value == "?":
        fieldName = startName
    return fieldName, value

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

def formateURL(url, regexpSite):
    """ Formate une URL en HTML """
    url = url.strip()
    tagA = url
    match = regexpSite.search(url)
    if match:
        tagA = '<a href="' + url + '" target="_blank">'
        tagA += match.group('id') + ' (' + match.group('siteName') + ')'
        tagA += '</a>'
    return tagA

def genKMLFiles(listInfoRead, titleKML, URLPicto, pathKMLFile, isVerbose):
    """ genere les fichiers de sortie KML"""

    # Ref simplekml : http://www.simplekml.com/en/latest/reference.html
    import simplekml

    withPicto = (len(URLPicto) > 0)

    if withPicto and not URLPicto.startswith("http"):
        raise ValueError("URL du picto incorrecte :" +
                         URLPicto +
                         " : devrait commencé par http")

    print("Ecriture des résultats dans", pathKMLFile, "...")
    titleKML = titleKML + " " + time.strftime("%d/%m/%y")
    kml = simplekml.Kml(name=titleKML)

    # Style icone et couleur du texte pour tous les éléments
    # Ref couleur : http://www.simplekml.com/en/latest/constants.html#color
    styleIcon = simplekml.Style()
    styleIcon.labelstyle.color = simplekml.Color.cadetblue
    if withPicto:
        styleIcon.iconstyle.icon.href = URLPicto

    for element in listInfoRead:
        point = kml.newpoint(name=element['nom'],
                             description='<![CDATA[' + element['description'] + ']]>\n',
                             coords=[(str(element['longitude']), str(element['latitude']))])
        point.style = styleIcon

    kml.save(pathKMLFile)
    print(str(len(listInfoRead)), "éléments écrits dans", pathKMLFile)

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

        self.URL_PICTO_DEFAULT = "https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png"

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

        # Pour affichage des éléments lus
        elementsFrame = tkinter.LabelFrame(mainFrame, text="elements trouvés dans le fichier")
        self.elementsListbox = tkinter.Listbox(elementsFrame,
                                       background="light blue",
                                       height=10, width=70)
        self.elementsListbox.grid(row=0, columnspan=2)
        scrollbarRightelements = tkinter.Scrollbar(elementsFrame, orient=tkinter.VERTICAL,
                                                  command=self.elementsListbox.yview)
        scrollbarRightelements.grid(row=0, column=2, sticky=tkinter.W+tkinter.N+tkinter.S)
        self.elementsListbox.config(yscrollcommand=scrollbarRightelements.set)
        scrollbarBottomelements = tkinter.Scrollbar(elementsFrame, orient=tkinter.HORIZONTAL,
                                                   command=self.elementsListbox.xview)
        scrollbarBottomelements.grid(row=1, columnspan=2, sticky=tkinter.N+tkinter.E+tkinter.W)
        self.elementsListbox.config(xscrollcommand=scrollbarBottomelements.set)
        elementsFrame.pack(side = tkinter.TOP, fill="both", expand="yes")

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
        self.urlPictoVar.set(self.URL_PICTO_DEFAULT)

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
                raise ValueError("Aucun élément trouvé dans le fichier !")

            # Update message list
            self.messagesListbox.delete(0, tkinter.END)
            for message in self.listMessage:
                self.messagesListbox.insert(tkinter.END,
                                            "Ligne " + str(message['numLigne']) +
                                            " : " + message['texte'])
            # Update élément list
            self.elementsListbox.delete(0, tkinter.END)
            for element in self.listInfoRead:
                self.elementsListbox.insert(tkinter.END,
                                           element['Commune'] + " : " + element['nom'])

            self.setMessageLabel("Fichier converti en .kml : " + str(len(self.listMessage)) +
                                        " messages, " + str(len(self.listInfoRead)) +
                                        " elements lus dans : " +
                                        os.path.basename(pathFicTable))
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
