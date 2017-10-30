#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
"""
*********************************************************
Programme : dolmenxls2kml.py
Github : https://github.com/Thierry46/dolmenxls2kml
Auteur : Thierry Maillard (TMD)
Date : 27 - 30/10/2017

Role:Convertir les coordonées informations de localisation contenues
dans un fichier Excel au format KML importable dans Geoportail.
Utilise le package simplekml pour écrire le fichier au format kml.

Prerequis :
- Python v3.xxx : a télécharger depuis : https://www.python.org/downloads/
- module xlrd : sudo python3 -m pip install xlrd
- module simplekml : sudo python3 -m pip install simplekml

Usage : dolmenxls2kml.py [-h] [-v] [Chemin xls]
Sans paramètre, lance une IHM, sinon fonctionne en batch avec 1 parametre.

Environnement :
    Ce programme teste son environnement (modules python disponibles)
    et s'y adaptera.
    tkinter : pour IHM : facultatif (mode batch alors seul)
    xlrd : pour lire le fichier  Excel (obligatoire)
    simplekml : pour ecrire le fichier resultat kml (obligatoire)

Parametres :
    -h ou --help:affiche cette aide.
    -v ou --verbose:mode bavard
    Nom d'un fichier de configuration Excel (mode batch)

Sortie :
    - Fichier de même nom que .xls mais avec extension .kml

Modifications :
v0.2 : utilisation package simplekml + champ description + picto dolmen.
v0.3 : IHM tkinter
v0.4 : Decodage colonne Etat du fichier .xls
v0.5 : Add external links in KML Description tag

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
import threading
import re

##################################################
# main function
##################################################
def main(argv=None):
    VERSION = 'v0.5 - 30/10/2017'
    NOM_PROG = 'dolmenxls2kml.py'
    isVerbose = False
    title = NOM_PROG + ' - ' + VERSION + " sur " + platform.system() + " " + platform.release()
    print(title)

    # Test environnement
    # Test presence des modules tkinter, numpy et matplotlib
    canUseGUI = True
    try :
        imp.find_module('tkinter')
        import tkinter
        print("Module tkinter OK.")
    except ImportError as exc:
        print("Warning ! module tkinter pas disponible -> utiliser mode batch :", exc)
        canUseGUI = False
    try :
        imp.find_module('xlrd')
        print("Module xlrd OK.")
    except ImportError as exc:
        print(__doc__)
        print("Erreur ! module xlrd pas disponible :", exc)
        sys.exit(1)
    try :
        imp.find_module('simplekml')
        print("Module simplekml OK.")
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
            print("Mode verbose :b avard pour debug")

    if len(args) < 1:
        if canUseGUI:
            import tkinter
            print("Lancement de l'IHM...")
            root = tkinter.Tk()
            dolmenxls2kmlGUI(root, dirProject, title, isVerbose)
            root.mainloop()
        else:
            print(__doc__)
            print("IHM non disponible : tkinter pas installe !")
            print("Utilisez le mode batch et passez le fichier Excel au programme")
            sys.exit(2)

    else: # Batch
        processFile(args[0], isVerbose)

    print('End dolmenxls2kml.py', VERSION)
    sys.exit(0)

def processFile(pathFicExcel, isVerbose):
    """ Convertit un fichier Excel passé en paramètre en un fichier KML """
    listMessage, listDolmen = readExcel(pathFicExcel, isVerbose)
    pathKMLFile = pathFicExcel.replace(".xls", ".kml")
    genKMLFiles(listDolmen, pathKMLFile, isVerbose)
    return listMessage, listDolmen

def readExcel(pathFicExcel, isVerbose):
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
        'Etat' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Tumulus' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Orthostats' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Table' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Classement' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'Détails' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'URL' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None},
        'OSM' : {'obligatoire':False, 'numCol':-1, 'nomCol':"", 'dataCol':None}
            }

    if not pathFicExcel.endswith(EXT_FIC_OK):
        raise ValueError("Nom configuration incorrect:" +
                          os.path.basename(pathFicExcel) +
                          ":devrait finir par l'extension .xls")

    print("Lecture de", pathFicExcel, "...")
    # ouverture du fichier Excel 
    workbook = xlrd.open_workbook(filename= pathFicExcel, on_demand=True)
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
                value.strip() for value in sheetData.col_values(dictData[startColumn]['numCol'])]

    # Formatage et contrôle des donnees utiles
    listDolmen = []
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
        for field in ('Commune', 'Lieu', 'Lat', 'Lon', 'Classement', 'Détails', 'Etat',
                      'Tumulus', 'Orthostats', 'Table', 'OSM', 'URL'):
            if ligneOK and dictData[field]['numCol'] != -1 and \
                numLigne < len(dictData[field]['dataCol']) and \
                len(dictData[field]['dataCol'][numLigne]) != 0 :
                description += "<b>" + dictData[field]['nomCol'] + "</b> : "

                # Champs particuliers
                if field == 'Etat':
                    description += '\n<ul>\n'
                    codeEtat = dictData[field]['dataCol'][numLigne]
                    if '(C)' in codeEtat:
                        description += '<li>* Dalle de couverture présente, mais déplacée</li>\n'
                    elif 'C' in codeEtat:
                        description += '<li>* Dalle de couverture présente</li>\n'

                    if '1O' in codeEtat:
                        description += '<li>* Un seul orthostat</li>\n'
                    elif 'O' in codeEtat:
                        description += '<li>* Orthostats présents</li>\n'

                    if 'T' in codeEtat:
                        description += '<li>* Tumulus présent</li>\n'
                    if 'ro' in codeEtat:
                        description += "<li>* Restes d'orthostats</li>\n"
                    if 't' in codeEtat or 'c' in codeEtat :
                        description += '<li>* Restes de table</li>\n'
                    if '?' in codeEtat :
                        description += '<li>* Indéterminé</li>\n'
                    description += '</ul>\n'

                elif field == 'Commune':
                    nomCommune = dictData[field]['dataCol'][numLigne]
                    url = 'https://fr.wikipedia.org/wiki/' + nomCommune
                    description += '<a href="' + url + '" target="_blank">' + nomCommune + \
                                   '</href><br/>\n'

                elif field == 'URL':
                    url = dictData[field]['dataCol'][numLigne]
                    description += '<a href="' + url + '" target="_blank">Infos WEB</href><br/>\n'

                else:
                    description += dictData[field]['dataCol'][numLigne]
                    description += '<br/>\n'

        # Enregistrement des valeurs utiles dans la structure résultat
        if ligneOK:
            listDolmen.append({'numLigne':numLigne+1, 'nom':nomDolmen,
                              'Commune' : dictData['Commune']['dataCol'][numLigne],
                              'latitude':coordValue['Lat'],
                              'longitude':coordValue['Lon'],
                              'description':description
                              })
        else:
            listMessage.append(messageInfos)

    print(len(listDolmen), "dolmens enregistrés,", len(listMessage), "lignes ignorées.")

    if isVerbose:
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])

    return listMessage, listDolmen

def convertCoord(coord):
    """ Convertit une chaine de coordonnées d'angle en un réel
        La chaine d'entrée peut contenir une valeur sexagésimale """
    valFloat = -1.0
    try:
        valFloat = float(coord)
    except ValueError:
        # re OK pour expression du type 44°51'37" ou 1°51'37"
        m = re.search(r'(?P<degres>[\d]{1,2})°(?P<minutes>[\d]{2})\'(?P<secondes>[\d]{2})"$',
                      coord)
        if m:
            valFloat = float(m.group('degres')) + float(m.group('minutes')) / 60. + \
                       float(m.group('secondes')) / 3600.
        else:
            raise
    return valFloat

def genKMLFiles(listDolmen, pathKMLFile, isVerbose):
    """ genere les fichiers de sortie KML"""

    import simplekml
    # Ref simplekml:http://www.simplekml.com/en/latest/reference.html
    ICONE_FILE_URL = "https://upload.wikimedia.org/wikipedia/commons/e/eb/PointDolmen.png"

    print("Ecriture des résultats dans", pathKMLFile, "...")
    kml = simplekml.Kml(name='Dolmens Adrien')

    # Style icone et couleur du texte pour tous les dolmens
    # Ref couleur:http://www.simplekml.com/en/latest/constants.html#color
    styleIconDolmen = simplekml.Style()
    styleIconDolmen.labelstyle.color = simplekml.Color.cadetblue
    styleIconDolmen.iconstyle.icon.href = ICONE_FILE_URL

    for dolmen in listDolmen:
        point = kml.newpoint(name=dolmen['nom'],
                             description='<![CDATA[' + dolmen['description'] + ']]>\n',
                             coords=[(str(dolmen['longitude']), str(dolmen['latitude']))])
        point.style = styleIconDolmen

    kml.save(pathKMLFile)
    print(str(len(listDolmen)), "dolmens écrits dans", pathKMLFile)

############
class dolmenxls2kmlGUI():
    """
    A GUI for dolmenxls2kml script.
    """
    def __init__(self, root, dirProject, title, isVerbose):
        """
        Constructor
        Define all GUIs widgets
        parameters :
            - root : main window
            - title : Windows title
            - isVerbose : If true, environment is OK for plotting
        """
        import tkinter
        self.root = root
        mainFrame = tkinter.Frame(self.root)
        self.dirProject = dirProject
        self.isVerbose = isVerbose
        self.listMessage = []
        self.listDolmen = []

        self.root.title(title)

        # Frame des paramètres de configuration
        inputFrame = tkinter.LabelFrame(mainFrame, text="Parametres")
        self.readConfigButton = tkinter.Button(inputFrame,
            text="Fichier Excel ...",
            command=self.fileChooserExcelCallback,
            bg='green')
        self.readConfigButton.grid(row=0, column=0)
        self.pathExcelVar = tkinter.StringVar()
        pathConfigurationExcel = tkinter.Entry(inputFrame,
                                          textvariable=self.pathExcelVar,
                                          width=60)
        pathConfigurationExcel.grid(row=0, column=1, sticky=tkinter.W)
        pathConfigurationExcel.bind('<Return>', self.launchInputFileReader)
        pathConfigurationExcel.bind('<KP_Enter>', self.launchInputFileReader)
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
        dolmensFrame = tkinter.LabelFrame(mainFrame, text="Dolmen trouvés dans le fichier Excel")
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
        
    def fileChooserExcelCallback(self) :
        """
        Callback for button that launch file chooser for Excel file.
        """
        from tkinter import filedialog

        # Le dernier répertoire du fichier de travail est relu pour
        #        faciliter les traitements répétitifs.
        lastDir = "."
        try :
            with open(os.path.join(self.dirProject, "dolmenxls2kmlDir.txt"), 'r') as hDirFile:
                lastDir = hDirFile.readline().strip()
                hDirFile.close()      
        except IOError:
            pass
        fileName = filedialog.askopenfilename(
                        parent=self.root,
                        initialdir=lastDir,
                        filetypes = [("Fichier Excel","*.xls"), ("All", "*")],
                        title="Selectionnez un fichier de configuration")
        if fileName != "":
            # Le dernier répertoire du fichier de travail est stocké pour
            #        faciliter les traitements répétitifs.
            try :
                with open(os.path.join(self.dirProject, "dolmenxls2kmlDir.txt"), 'w') as hDirFile:
                    hDirFile.write(os.path.dirname(fileName) + "\n")
                    hDirFile.close()      
            except IOError:
                self.setMessageLabel(str(exc), True)
            self.pathExcelVar.set(fileName)
            self.launchInputFileReader()

    def launchInputFileReader(self, event=None):
        """ Update list according input files
            v2.0:Thread use"""
        self.setMessageLabel("Lecture du fichier Excel en cours, veuillez patienter...")
        import tkinter

        pathFicExcel = self.pathExcelVar.get()
        try :
            self.listMessage, self.listDolmen = \
                processFile(pathFicExcel, self.isVerbose)
            if len(self.listDolmen) == 0:
                raise ValueError("Aucun dolmen trouvé dans le fichier Excel !")

            # Update message list
            self.messagesListbox.delete(0, tkinter.END)
            for message in self.listMessage:
                self.messagesListbox.insert(tkinter.END,
                                            "Ligne " + str(message['numLigne']) +
                                            " : " + message['texte'])
            # Update dolmen list
            self.dolmensListbox.delete(0, tkinter.END)
            for dolmen in self.listDolmen:
                self.dolmensListbox.insert(tkinter.END,
                                           dolmen['Commune'] + " : " + dolmen['nom'])

            self.setMessageLabel("Fichier converti : " + str(len(self.listMessage)) +
                                        " messages, " + str(len(self.listDolmen)) +
                                        " dolmens dans : " +
                                        os.path.basename(pathFicExcel).replace('.xls', '.kml'))
        except IOError as exc:
            self.setMessageLabel(str(exc), isError=True)
        except OSError as exc:
            self.setMessageLabel(str(exc), isError=True)
        except ValueError as exc:
            self.setMessageLabel(str(exc), isError=True)

    def setMessageLabel(self, message, isError=False) :
        """ set a message for user. """
        if isError:
            message = "Probleme:" + message
        self.messageLabel['text'] =  message
        if isError:
            print(message)

#to be called as a script:python dolmenxls2kml.py or dolmenxls2kml.py
if __name__ == "__main__":
    main()
