#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
"""
    *********************************************************
    Programme : taisne2cvs.py
    Github : https://github.com/Thierry46/table2kml
    Auteur : Thierry Maillard (TMD)
    Date : 11/11/2017

    Role : Extrait les informations depuis le Taisne transformé
    du format Pdf au format CSV et les écrit dans un fichier CSV
    utilisatble par tabele2kml en vue de produire un calque
    Geoportail au format KML.

    Prerequis :
    - Python v3.xxx : a télécharger depuis : https://www.python.org/downloads/
    - Livre de Jean Taisne au format pdf convertit en .txt avec l'outil
    pdfminer :
    sudo python -m pip install pdfminer
    pdf2txt.py -o taisne.txt taisne.pdf 795 Ko

    Usage : taisne2cvs.py [-h] [-v] taisne.txt

    Parametres :
    -h ou --help : affiche cette aide.
    -v ou --verbose : mode bavard
    Nom d'un fichier de données .txt

    Sortie :
    - Fichier de même nom que fichier d'entrée mais avec extension .csv

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
import platform
import getopt
import os.path
import math
import imp
import re
import csv

_PREC_COORD_DEC_ = 6

##################################################
# main function
##################################################
def main(argv=None):
    VERSION = 'v1.4 - 18/11/2017'
    NOM_PROG = 'taisne2cvs.py'
    isVerbose = False
    title = NOM_PROG + ' - ' + VERSION + " sur " + platform.system() + " " + platform.release() + \
        " - Python : " + platform.python_version()
    print(title)

    try :
        imp.find_module('simplekml')
        print("Module pyproj OK : Conversion coordonnées Lambert3 -> WGS84 autorisé")
    except ImportError as exc:
        print(__doc__)
        print("Erreur ! module pyproj pas disponible :", exc)
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

    if len(args) == 1:
        processFile(args[0], isVerbose)
    else:
        print(__doc__)
        print("Nombre de paramètre invalide : 1 nécessaires : chemin fichier Taisne")
        sys.exit(1)

    print('End', NOM_PROG, VERSION)
    sys.exit(0)

def processFile(pathFicTaisne, isVerbose):
    """ Extrait les infos du fichier taisne passé en paramètre
        et produit un fichier .csv
        """
    if not pathFicTaisne.endswith(".txt"):
        raise ValueError("Extension du fichier non supporté :" +
                         os.path.basename(pathFicTaisne) +
                         " extension supportées : .txt")

    print("Conversion du fichier", pathFicTaisne, "...")
    pathFicTaisneCSV = pathFicTaisne.replace(".txt", ".csv")
    nbCaviteOK = 0
    with open(pathFicTaisneCSV, 'w', newline='') as hFicCSV:
        writer = csv.writer(hFicCSV, delimiter=';', quoting=csv.QUOTE_ALL)
        columnTitle = ['Nom cavité', 'Alias', 'Commune', 'IGN',
                       'X Lambert3', 'Y Lambert3', 'Latitude', 'Longitude',
                       'Altitude', 'Description', 'Page Taisne', 'Plan']
        writer.writerow(columnTitle)

        # Regexp to analyse parts
        regexpTitle = re.compile(r'^(?P<titre>INVENTAIRE ALPHABÉTIQUE)$')
        regexpPage = re.compile(r'^(?P<page>[\d]+)$')
        regexpCaveName = re.compile(r"^(?P<caveName>[A-Z0-9 \-°ÏÑÉÂÈÇËÈÊÜÔÛŒô\?\(\)\'\’]+?), " +
                                    r'(?P<startCaveName>.+)Commune d.(?P<commune>.*)')
        regexpVoirA =  re.compile(r'- voir à')
        regexpCoord = re.compile(r'^(?P<Xe>\d{3}),(?P<Xd>\d{2}) - (?P<Ye>\d{3}),(?P<Yd>\d{2})' +
                                 r' - (?P<altitude>\d{3})m[\( \)]*IGN (?P<IGN>\d{4} [ETOW]{1,2})\)$')
        regexpCoordMult = re.compile(r'(?P<sousGrotte>.*?)[: ]?' +
                                     r'(?P<Xe>\d{3}),(?P<Xd>\d{2}) - (?P<Ye>\d{3}),(?P<Yd>\d{2})' +
                                     r' - (?P<altitude>\d{3})m[ \)]?$')
        regexpCoordMultIGN = re.compile(r'(?P<sousGrotte>.*?)[: ]?' +
                                        r'(?P<Xe>\d{3}),(?P<Xd>\d{2}) - (?P<Ye>\d{3}),(?P<Yd>\d{2})' +
                                        r' - (?P<altitude>\d{3})m \) ' +
                                        r'\(IGN (?P<IGN>\d{4} [ETOW]{1,2})\)$')
        regexpPlan = re.compile(r'Plan (?P<plan>\d{1,3}).')
        regexpIGNSeul = re.compile(r'^[\) \(]?IGN (?P<IGN>\d{4} [ETOW]{1,2})\)$')

        caveOk = False
        wait4Coord = False
        wait4Description = False
        numPage = 0
        caveName = ""
        startCaveName = ""
        alias = ""
        IGN = ""
        description = ""
        plan = ""
        listeCoordEntree = []
        with open(pathFicTaisne, 'r') as hFicTaisne:
            for numLine, line in enumerate(hFicTaisne.read().splitlines()):
                ignoreLine = False
                error = False

                # Titre inventaire
                match = regexpTitle.search(line)
                if match and numLine == 0:
                    if isVerbose:
                        print('ligne ', numLine+1, ' : titre OK : ', match.group('titre'))
                elif match and numLine != 0:
                    message = 'titre sur mauvaise ligne : ' + match.group('titre')
                    error = True

                # Elimination des renvois
                if not match and not error:
                    match = regexpVoirA.search(line)

                # Numéro de page
                if not match and not error:
                    match = regexpPage.search(line)
                    if match:
                        numPageLu = int(match.group('page'))
                        if numPageLu > numPage : #Pb chaine CO2 mal lue par pdfminer
                            numPage = numPageLu
                            if isVerbose:
                                print(numLine+1, ': Page : ', numPage)

                # ligne nom cavité
                if not match and not error:
                    match = regexpCaveName.search(line)
                    if match:
                        if caveOk:
                            writeCave(writer, caveName, startCaveName, alias, commune, IGN,
                                      listeCoordEntree, description, numPage, plan, isVerbose)
                            nbCaviteOK += 1
                            caveOk = False
                            wait4Description = False
                            caveName = ""
                            startCaveName = ""
                            alias = ""
                            listeCoordEntree = []
                            IGN = ""
                            description = ""
                            plan = ""

                        caveName = match.group('caveName')
                        startCaveName = match.group('startCaveName')
                        startCaveName = startCaveName[:-1*len(' - ')]
                        if '(nE' in startCaveName:
                            startCaveName = startCaveName.replace('(nE', 'n°').replace(')', '')
                        aliasStart = startCaveName.find(' - ')
                        if aliasStart != -1:
                            alias = startCaveName[aliasStart + len(' - '):]
                            startCaveName = startCaveName[:aliasStart]
                        else:
                            alias = ""
                        commune = match.group('commune')
                        if commune.startswith(' '):
                            commune = commune[1:]
                        wait4Coord = True
                        if isVerbose:
                            print(numLine+1, ': qualif :', startCaveName, 'nom :', caveName,
                                  'alias :', alias,
                                  'commune :', commune)

                # Ligne coordonnées complete
                if not match and not error:
                    match = regexpCoord.search(line)
                    if match:
                        if not wait4Coord:
                            message = 'coordonnée sur mauvaise ligne : '
                            error = True
                        else:
                            xLambert3, yLambert3, longitude, latitude = \
                                    parseCoordinates(match, numLine, isVerbose)
                            listeCoordEntree = [{'nom':"",
                                                'xLambert3':xLambert3,
                                                'yLambert3':yLambert3,
                                                'longitude':longitude,
                                                'latitude':latitude,
                                                'altitude':match.group('altitude')
                                                }]

                            IGN = match.group('IGN')
                            wait4Coord = True
                            wait4Description = True
                            if isVerbose:
                                print('IGN :', IGN)

                # Ligne coordonnes sous grotte
                if not match and not error:
                    match = regexpCoordMult.search(line)
                    if match:
                        if not wait4Coord:
                            message = 'coordonnée sous grotte sur mauvaise ligne : '
                            error = True
                        else:
                            xLambert3, yLambert3, longitude, latitude = \
                                parseCoordinates(match, numLine, isVerbose)
                            listeCoordEntree.append({'nom':match.group('sousGrotte'),
                                                'xLambert3':xLambert3,
                                                'yLambert3':yLambert3,
                                                'longitude':longitude,
                                                'latitude':latitude,
                                                'altitude':match.group('altitude')
                                                })

                            wait4Coord = True
                            wait4Description = True

                # Ligne coordonnes sous grotte
                if not match and not error:
                    match = regexpCoordMultIGN.search(line)
                    if match:
                        if not wait4Coord:
                            message = 'coordonnée sous grotte sur mauvaise ligne : '
                            error = True
                        else:
                            xLambert3, yLambert3, longitude, latitude = \
                                parseCoordinates(match, numLine, isVerbose)
                            listeCoordEntree.append({'nom':match.group('sousGrotte'),
                                                    'xLambert3':xLambert3,
                                                    'yLambert3':yLambert3,
                                                    'longitude':longitude,
                                                    'latitude':latitude,
                                                    'altitude':match.group('altitude')
                                                    })
                            IGN = match.group('IGN')
                            if isVerbose:
                                print('IGN :', IGN)
                            wait4Coord = True
                            wait4Description = True

                # Ligne IGN isolée
                if not match and not error:
                    match = regexpIGNSeul.search(line)
                    if match:
                        IGN = match.group('IGN')
                        wait4Coord = False
                        wait4Description = True
                        if isVerbose:
                            print('IGN seul :', IGN)

                # Lignes description
                if not match and not error and wait4Description:
                    if line.strip() != "":
                        description += line.strip() + '<br/>\n'
                        caveOk = True
                        wait4Coord = False
                        matchPlan = regexpPlan.search(line)
                        if matchPlan:
                            plan = matchPlan.group('plan')
                            if isVerbose:
                                print(numLine+1, ': Plan : ', plan)

                if error:
                    print('ligne ', numLine+1, ' : Erreur : ', message, ' :', line)
                    error = False

            if caveOk:
                writeCave(writer, caveName, startCaveName, alias, commune, IGN,
                          listeCoordEntree, description, numPage, plan, isVerbose)
                nbCaviteOK += 1
            print("Nombre de cavité extraites :", nbCaviteOK)

def parseCoordinates(match, numLine, isVerbose):
    xLambert3 = float(match.group('Xe')) + float(match.group('Xd')) / 100.
    yLambert3 = 3000. + float(match.group('Ye')) + float(match.group('Yd')) / 100.
    longitude, latitude = lambert3ToWGS84(xLambert3 * 1000., yLambert3 * 1000.)
    altitude = float(match.group('altitude'))
    if isVerbose:
        print(numLine+1,
              'Lambert3 :', xLambert3, yLambert3,
              'WGS84 :', longitude, latitude)
        print('Altitude :', altitude)
    return xLambert3, yLambert3, longitude, latitude

def writeCave(writer, caveName, startCaveName, alias, commune, IGN,
              listeCoordEntree, description, numPage, plan, isVerbose):
    # Ecrit la cavite précedente
    nom = startCaveName
    if not startCaveName.endswith("'"):
        nom += " "
    nom += caveName

    if description.endswith('<br/>\n'):
        description = description[:-1*len('<br/>\n')]
        if isVerbose:
            print('Description : ', description)

    for entree in listeCoordEntree:
        nomEntree = nom
        if entree['nom'] != "":
            nomEntree += ' : ' + entree['nom']
        writer.writerow([nomEntree, alias, commune, IGN,
                     entree['xLambert3'], entree['yLambert3'],
                     entree['latitude'], entree['longitude'],
                     entree['altitude'], description, numPage, plan])


def lambert3ToWGS84(xLambert3, yLambert3):
    """ Convertit des coordonnées Lambert3 : X (m), Y (m)
        ex. pour perte de l'abois à Assier : 564500, 3263475
        En coordonnées géographiques WGS84 longitude E, latitude N (degrés décimaux)
        Ref : https://rcomman.de/conversion-de-coordonnees-geographiques-en-python.html
    """
    import pyproj
    wgs84 = pyproj.Proj('+proj=longlat +ellps=WGS84 +datum=WGS84 +no_defs')
    lambert = pyproj.Proj('+proj=lcc +nadgrids=ntf_r93.gsb,null +towgs84=-168.0000,-60.0000,320.0000 +a=6378249.2000 +rf=293.4660210000000 +pm=2.337229167 +lat_0=44.100000000 +lon_0=0.000000000 +k_0=0.99987750 +lat_1=44.100000000 +x_0=600000.000 +y_0=3200000.000 +units=m +no_defs')
    longitude, latitude = pyproj.transform(lambert, wgs84, xLambert3, yLambert3)
    return round(longitude, _PREC_COORD_DEC_), round(latitude, _PREC_COORD_DEC_)

#to be called as a script:python taisne2cvs.py or taisne2cvs.py
if __name__ == "__main__":
    main()

