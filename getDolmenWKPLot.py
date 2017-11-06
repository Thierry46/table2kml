#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
"""
*********************************************************
Programme : getDolmenWKPLot.py
Github : https://github.com/Thierry46/table2kml
Auteur : Thierry Maillard (TMD)
Date : 5 - 6/11/2017

Role : Extrait de Wikipedia la liste des dolmens du Lot et
        leurs coordonnées.

Prerequis :
- Python v3.xxx : a télécharger depuis : https://www.python.org/downloads/

Usage : getDolmenWKPLot.py [-h] [-v]
Fonctionne en batch avec 1 parametre.

Parametres :
    -h ou --help : affiche cette aide.
    -v ou --isVerbose : mode bavard

Sortie :
- Fichier .csv compatible avec le programme de conversion csv -> kml : table2kml

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

import urllib.request, urllib.error, urllib.parse

##################################################
# main function
##################################################
def main(argv=None):
    VERSION = 'v1.0 - 6/11/2017'
    NOM_PROG = 'getDolmenWKPLot.py'
    NOM_ARTICLE_WIKIPEDIA = 'Sites mégalithiques du Lot'
    isVerbose = False
    title = NOM_PROG + ' - ' + VERSION + " sur " + platform.system() + " " + platform.release() + \
            " - Python : " + platform.python_version()
    print(title)

    if argv is None:
        argv = sys.argv

    # parse command line options
    dirProject = os.path.dirname(os.path.abspath(sys.argv[0]))
    try:
        opts, args = getopt.getopt(argv[1:], "hv", ["help", "isVerbose"])
    except getopt.error as msg:
        print(msg)
        print("To get help use --help ou -h")
        sys.exit(1)
    # process options
    for o, a in opts:
        if o in ("-h", "--help"):
            print(__doc__)
            sys.exit(0)

        if o in ("-v", "--isVerbose"):
            isVerbose = True
            print("Mode isVerbose : bavard pour debug")

    if len(args) != 0:
        print(__doc__)
        print("Aucun paramètre utilisé !")
        sys.exit(2)

    try:
        columnTitleMap, listInfoReadMap = getInfoFromWikipedia(NOM_ARTICLE_WIKIPEDIA, isVerbose)
        writeCSV(columnTitleMap, listInfoReadMap, "carte")
    except ValueError as exc:
        print(str(exc))

    print('End getDolmenWKPLot.py', VERSION)
    sys.exit(0)

def getInfoFromWikipedia(nomArticleWikipedia, isVerbose):
    """ Extrait les info sur les dolmens de la page Wikipedia du Lot """
    listInfoReadMap = []
    columnTitleMap = []
    listMessage = []

    # Syntaxe des expressions régulières utilisées :
    # \x : caractère x qui est normalement un caractère spécial
    # \d* : plusieurs chiffres
    # ^ : Début de ligne

    # pour recherche chaine du genre :
    #'{{G|Lot|44.84015|1.68073|Dolmen de Viroulou {{n°|1}}|Grotte sans toponyme}}'
    regexpMap = re.compile(r'^[ ]*\{\{G\|Lot\|[ ]*(?P<Lat>\d*.\d*)[ ]*\|[ ]*(?P<Lon>\d*\.\d*)[ ]*' +
                           r'\|(?P<Nom>.*)\|.*sans toponyme\}\}')

    print("Recup des dolmens de l'article :", nomArticleWikipedia, "...")
    nomArticleUrl = urllib.request.pathname2url(nomArticleWikipedia)
    page = getPageWikipediaFr(nomArticleUrl, isVerbose)

    # Lecture des dolmens de la carte
    columnTitleMap = ['Nom', 'Lat', 'Lon', 'Commune']
    inCarte = False
    for numRow, line in enumerate(page.splitlines()):
        if not inCarte and '{{Début de carte}}' in line:
            inCarte = True
        elif inCarte and '{{Fin de carte}}' in line:
            inCarte = False
        elif inCarte:
            match = regexpMap.search(line)
            if match:
                nom = remove_chars(match.group('Nom'), "{}|")
                listInfoReadMap.append([nom, match.group('Lat'), match.group('Lon'), "?"])
            else:
                messageInfos = "Ligne ignorée dans section carte : " + line
                listMessage.append({'numLigne':numRow, 'texte':messageInfos})

    print(len(listInfoReadMap), "dolmens lus dans la section carte et enregistrés")
    if isVerbose:
        print("Anomalies détectées dans la section carte :")
        print("-----------------------------------------")
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])
        print("-----------------------------------------")
    return columnTitleMap, listInfoReadMap

def remove_chars(subj, chars):
    sc = set(chars)
    return ''.join([c for c in subj if c not in sc])

def writeCSV(columnTitle, listInfoRead, type):
    """ Ecrit les informations dans le fichier CSV dans un format compatible avec table2kml """
    import csv
    if len(listInfoRead) == 0:
        raise ValueError("Aucun dolmen à écrire !")

    titleCSVFile = "wikipedia_fr_" + type + "_" + time.strftime("%Y_%m_%d") + ".csv"
    print("Ecriture des résultats dans", titleCSVFile, "...")
    with open(titleCSVFile, 'w', newline='') as hFicCSV:
        writer = csv.writer(hFicCSV)
        writer.writerow(columnTitle)
        for dolmen in listInfoRead:
            writer.writerow(dolmen)

def getPageWikipediaFr(nomArticleUrl, isVerbose):
    """
        Ouvre une page de Wikipedia et retourne le texte brut de la page
        if problem with urllib ssl.SSLError :
        Launch "Install Certificates.command" located in Python installation directory :
        sudo /Applications/Python\ 3.6/Install\ Certificates.command
    """
    if isVerbose:
        print("Entrée dans getPageWikipediaFr")
        print("Recuperation de l'article :", nomArticleUrl)

    # Pour ressembler à un navigateur Mozilla/5.0
    opener = urllib.request.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]

    # Construction URL à partir de la config et du nom d'article
    baseWkpFrUrl = 'https://fr.wikipedia.org/wiki/'
    actionTodo = '?action=raw'
    urltoGet = baseWkpFrUrl + nomArticleUrl + actionTodo
    if isVerbose:
        print("urltoGet =", urltoGet)

    # Envoi requete, lecture de la page et decodage vers Unicode
    infile = opener.open(urltoGet)
    page = infile.read().decode('utf8')

    if isVerbose:
        print("Sortie de getPageWikipediaFr")
        print("Nombre de caracteres lus :", len(page))
    return page

#to be called as a script:python getDolmenWKPLot.py or getDolmenWKPLot.py
if __name__ == "__main__":
    main()
