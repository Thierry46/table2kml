#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
"""
*********************************************************
Programme : getDolmenWKPLot.py
Github : https://github.com/Thierry46/table2kml
Auteur : Thierry Maillard (TMD)
Date : 5 - 8/11/2017

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

# For performance : calculated once
# Regexp for [[lien|text]] or [[lien]]
__REGEXP_LINK_SIMPLE__ = re.compile(r'[\[]{2}(?P<link>.*?)[\]]{2}')
__REGEXP_LINK_ALIAS__ = re.compile(r'[\[]{2}(?P<link>.*?)\|.*?[\]]{2}')
# Regexp for before<ref>texte reference</ref>after
__REGEXP_REF_SIMPLE__ = re.compile(r'(?P<before>.*?)<ref>(?P<ref>.*?)</ref>(?P<after>.*)')
__REGEXP_REF_NAMED_DEF__ = re.compile(r'(?P<before>.*?)<ref name=.*?>(?P<ref>.*?)</ref>(?P<after>.*)')
__REGEXP_REF_NAMED_USED__ = re.compile(r'(?P<before>.*?)<ref name=(?P<ref>.*?)[ ]*/>(?P<after>.*)')
# regexp d'extraction des coordonnées
__REGEXP_COORD_DECIMAL__ = re.compile(r'[\{]{2}[ ]*[Cc]oord\|[ ]*(?P<Lat>\d*\.\d*)[ ]*\|' + \
                                  r'[ ]*(?P<Lon>\d*\.\d*)')
__REGEXP_COORD_SEXAGESIMAL__ = re.compile(r'[\{]{2}[ ]*[Cc]oord\|[ ]*(?P<Lat>\d*°' + \
                                          r"\d*'" + r'\d*")[ ]*N\|' + \
                                          r'[ ]*(?P<Lon>\d*°\d*'+ \
                                          r"\d*'" + r'\d*")[ ]*E')

##################################################
# main function
##################################################
def main(argv=None):
    VERSION = 'v2.0 - 8/11/2017'
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
        columnTitleMap, listInfoReadMap, columnTitleArticle, listInfoReadArticle = \
                getInfoFromWikipedia(NOM_ARTICLE_WIKIPEDIA, isVerbose)
        writeCSV(columnTitleMap, listInfoReadMap, "carte")
        writeCSV(columnTitleArticle, listInfoReadArticle, "liste")

    except ValueError as exc:
        print(str(exc))

    print('End getDolmenWKPLot.py', VERSION)
    sys.exit(0)

def getInfoFromWikipedia(nomArticleWikipedia, isVerbose):
    """ Extrait les info sur les dolmens de la page Wikipedia du Lot """

    # Syntaxe des expressions régulières utilisées :
    # \x : caractère x qui est normalement un caractère spécial
    # \d* : plusieurs chiffres
    # ^ : Début de ligne

    print("Recup des dolmens de l'article :", nomArticleWikipedia, "...")
    nomArticleUrl = urllib.request.pathname2url(nomArticleWikipedia)
    page = getPageWikipediaFr(nomArticleUrl, isVerbose)

    # Lecture des dolmens de la carte
    # pour recherche chaine du genre :
    #'{{G|Lot|44.84015|1.68073|Dolmen de Viroulou {{n°|1}}|Grotte sans toponyme}}'
    regexpMap = re.compile(r'^[ ]*\{\{G\|Lot\|[ ]*(?P<Lat>\d*\.\d*)[ ]*\|[ ]*(?P<Lon>\d*\.\d*)[ ]*' +
                           r'\|(?P<Nom>.*)\|.*sans toponyme\}\}')

    listInfoReadMap = []
    listMessage = []
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
        print("------------------------------------------")
        print(len(listMessage), "Anomalies détectées dans la section carte :")
        print("-----------------------------------------")
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])
        print("-----------------------------------------")

    ###########################################
    # Lecture des dolmens du corps de l'article
    ###########################################
    # Regroupement des lignes de table
    inListe = False
    inRow = False
    listRowContent = []
    rowContent = ""
    for numRow, line in enumerate(page.splitlines()):
        line = line.strip()
        errorRow = False
        if not inListe and '=== Liste non exhaustive ===' in line:
            inListe = True
        elif inListe and line.startswith('|}'):
            # Enregistre rowContent dans la liste
            if len(rowContent) != 0:
                listRowContent.append(rowContent)
            inListe = False
        elif inListe and not inRow and line.startswith('|-'):
            inRow = True
        elif inListe and inRow and not line.startswith('|-'):
            # Concat line with separator |
            if rowContent == "":
                rowContent = line
            else:
                rowContent += "|" + line
        elif inListe and inRow and line.startswith('|-'):
            # Enregistre rowContent dans la liste
            listRowContent.append(rowContent)
            rowContent = ""
    print(str(len(listRowContent)), "lignes groupe de dolmens trouvées dans la section Liste .")

    # pour recherche chaine du genre :
    #! Monument !! Commune !! Lieu !! Protection !! Localisation !! Image
    regexpListe = re.compile(r'^[ ]*\|[ ]*(?P<Nom>.*?)[ ]*[\|]{2}' + \
                             r'[ ]*[\[]{2}(?P<Commune>.*?)[\]]{2}.*?[\|]{2}' + \
                             r'[ ]*(?P<Lieu>.*?)[ ]*[\|]{2}' + \
                             r'[ ]*(?P<Protection>.*?)[ ]*[\|]{2}' + \
                             r'[ ]*(?P<Coordonnees>.*)[ ]*[\|]{2}')
    listInfoReadArticle = []
    listMessage = []
    columnTitleArticle = ['Nom', 'Commune', 'Pages Web', 'Remarques', 'Lieu', 'Classement',
                          'Lat', 'Lon']
    # Indice 1 pour ne pas prendre en compte la ligne de titre
    for numRow, line in enumerate(listRowContent[1:]):
        errorRow = False
        url = ""
        ref = ""
        lieu = ""
        classement = ""
        match = regexpListe.search(line)
        if match:
            # Traitement nom
            try:
                isLink, nom, ref = cleanNom(match.group('Nom'))
                if isLink:
                    url = 'https://fr.wikipedia.org/wiki/' + nom
            except ValueError as exc:
                errorRow = True
                messageRow = "Groupe dolmen : " + str(exc) + " pour champ nom : " + \
                    match.group('Nom')

            # Traitement commune
            if not errorRow:
                try:
                    isLink, commune = extractLink("[[" + match.group('Commune') + "]]")
                except ValueError as exc:
                    errorRow = True
                    messageRow = "Groupe dolmen : " + str(exc) + " pour champ Commune : " + \
                                 match.group('Commune')

            if not errorRow:
                # Traitement Lieu
                lieu = match.group('Lieu')

                # Traitement Lieu
                classement, refClassement = extractAllRef(match.group('Protection'))

                ref += " - " + refClassement

                # Recup coordonnées
                try:
                    listCoordDolmen = parseCoord(match.group('Coordonnees'))
                except ValueError as exc:
                    errorRow = True
                    messageRow = "Groupe dolmen : " + match.group('Nom') + ", " + str(exc)

            # enregistrement infos
            if not errorRow:
                nom = cleanField(nom)
                ref = cleanField(ref)
                lieu = cleanField(lieu)
                classement = cleanField(classement)
                # Enregistrement des dolmens membre de chaque groupe
                for dolmen in listCoordDolmen:
                    listInfoReadArticle.append([nom + " " + cleanField(dolmen['Nom']),
                                                commune, url, ref, lieu,
                                                classement, dolmen['Lat'], dolmen['Lon']])
            else:
                messageInfos = "Ligne ignorée " + messageRow + " : " + line
                listMessage.append({'numLigne':numRow, 'texte':messageInfos})
        else:
            messageInfos = "Ligne ignorée : " + line
            print("messageInfos=", messageInfos)
            listMessage.append({'numLigne':numRow, 'texte':messageInfos})

    print(len(listInfoReadArticle), "dolmens lus dans la section Liste et enregistrés")
    for grpDolmen in listInfoReadArticle:
        if 'Peyrelevade' in grpDolmen[0]:
            print(grpDolmen)
    if isVerbose:
        print("------------------------------------------")
        print(len(listMessage), "Anomalies détectées dans la section Liste :")
        print("------------------------------------------")
        for message in listMessage:
            print("Ligne numéro", message['numLigne'], message['texte'])
            print("-----------------------------------------")
    return columnTitleMap, listInfoReadMap, columnTitleArticle, listInfoReadArticle

def parseCoord(coordTxt):
    listCoordDolmen = []
    if len(coordTxt) == 0:
        raise ValueError("Champ coordonnées vide")
    if 'oord|' not in coordTxt:
        raise ValueError("Pas de coordonnées")

    # Dégrpage des dolmens
    listNomCoord = []
    if '<br>' in coordTxt:
        coordTxtSplit = coordTxt.split("<br>")
        if (len(coordTxtSplit) % 2) != 0:
            raise ValueError("Champ coordonnées manque une coordonnée ?")
        for numDolmen in range(0, len(coordTxtSplit), 2):
            if 'oord|' in coordTxtSplit[numDolmen]:
                listNomCoord.append(["Dolmen A", coordTxtSplit[numDolmen]])
                listNomCoord.append(["Dolmen B", coordTxtSplit[numDolmen+1]])
            else:
                listNomCoord.append([coordTxtSplit[numDolmen].strip(), coordTxtSplit[numDolmen+1]])
    else:
        listNomCoord.append(["", coordTxt])

    # extraction des coordonnées Longitude et latitude pour chaque dolmen
    for dolmen in listNomCoord:
        match = __REGEXP_COORD_DECIMAL__.search(dolmen[1])
        if not match:
            match = __REGEXP_COORD_SEXAGESIMAL__.search(dolmen[1])
        if match:
            latitude = match.group('Lat')
            longitude = match.group('Lon')
            listCoordDolmen.append({'Nom':dolmen[0], 'Lat':latitude, 'Lon':longitude})
        else:
            raise ValueError("Champ coordonnées Erreur format")
    return listCoordDolmen

def cleanNom(nomWkp):
    """ Suppress ref and expand model n° """
    nom, ref = extractAllRef(nomWkp)
    isLink, nom = extractLink(nom)
    # Expand model n°
    nom = remove_chars(nom, "{}|")
    return isLink, nom, ref

def extractAllRef(txtWkp):
    """ Extract all wikipedia reference """
    texte = txtWkp
    reftexteAll = ""
    otherRef = True
    while otherRef:
        texte, reftexte = extractRef(texte)
        if reftexte == "":
            otherRef = False
        else:
            reftexteAll += reftexte + "; "
    return texte, reftexteAll

def extractRef(txtWkp):
    """ Extract a wikipedia reference """
    texte = txtWkp.replace('{{,}}', '')
    reftexte = ""
    match = __REGEXP_REF_SIMPLE__.search(texte)
    if not match:
        match = __REGEXP_REF_NAMED_DEF__.search(texte)
    if not match:
        match = __REGEXP_REF_NAMED_USED__.search(texte)

    if match:
        texte = match.group('before') + match.group('after')
        reftexte = match.group('ref')

    return texte, reftexte

def extractLink(linkWkp):
    """ Extract a Wikipedia link from a field """
    islink = False
    resultStr = ""
    if linkWkp.startswith("[["):
        islink = True
        match = __REGEXP_LINK_ALIAS__.search(linkWkp)
        if not match:
            match = __REGEXP_LINK_SIMPLE__.search(linkWkp)
        if match:
            resultStr = match.group('link')
        else:
            raise ValueError("Problème lien incorrect : " + linkWkp)
    else:
        resultStr = linkWkp

    if len(resultStr) == 0:
        raise ValueError("Problème champ ou lien vide : " + linkWkp)
    return islink, resultStr

def remove_chars(subj, chars):
    sc = set(chars)
    return ''.join([c for c in subj if c not in sc])

def cleanField(field):
    field = remove_chars(field, '{}"')
    field = field.replace('|', ' ')
    field = field.replace("'",'')
    field = field.replace('harvsp ', '')
    field = field.replace('<br>', '<br/>')
    field = field.replace('opcit', '')
    field = field.replace('<ref name=', '')
    field = field.replace(' >', '')
    field = field.replace('</ref>', '')
    field = field.replace('/>', '')
    if field.startswith(" - "):
        field = field[len(" - "):]
    if field.endswith("; "):
        field = field[:-1*len("; ")]
    return field


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
