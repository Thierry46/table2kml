# table2kml
A tool to convert a table containing geographical information in KML Format.

Input file have to contain at least fields starting with : Nom, Lat, Lon
- Nom : Place name
- Lat : latitude WGS84
- Long : longitude WGS84

All column are displayed in description field except if their title begins with -

Input file format :
- .xls (Excel97) and .csv
- CSV file (prefered), all column are displayed in description field except if title begins with -

This simple tool is written in Python and uses tkinter, xlrd and simplekml packages.

It works in batch mode or with a GUI.

A program **getDolmenWKPLot.py** extracts datas from article Wikipedia [https://fr.wikipedia.org/wiki/Sites mégalithiques du Lot].

2 files are produced : one for the map and another for the list in article.
Outputs are in CSV format and can be converted with table2kml.

Pre-Requisites
-----------------

- [x] python3 :  [https://www.python.org/downloads] : Download python
- [ ] tkinter : usually installed with python : GUI toolkit, not needed in batch mode when you give a file on command line
- [x] simplekml : install : sudo python3 -m pip install simplekml : library used to write KML file
- [ ] xlrd : sudo python3 -m pip install xlrd : library used to read an Excel 97 file, not needed to convert .csv file

Installation
------------
* Copy table2kml.py on your computer
* Install Pre-Requisites
* Lauch the python file :
    * With GUI : python3 table2kml.py
    * In batch mode : python3 table2kml.py YOUR_FILE.xls ou YOUR_FILE.csv title URL_PICTO

Usage
-------
Launch *table2kml.py* with *-h* or *--help* option to get help or read comments at the beginning of the source file.

**./table2kml.py --help**

**./getDolmenWKPLot.py --help**

Example of output
------------------
[http://thierry.maillard.pagesperso-orange.fr/fr/dolmens/dolmens_lot.html]
![Dolmens sur carte du lot](http://thierry.maillard.pagesperso-orange.fr/fr/dolmens/screenshot_v0.5.jpg "Dolmens sur carte du lot")
![Info sur un dolment](http://thierry.maillard.pagesperso-orange.fr/fr/dolmens/screenshot_zoom_v0.5.jpg "Infos sur 1 dolmen")

Copyright
-----------
    Copyright 2017 Thierry Maillard
    This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
    You should have received a copy of the GNU General Public License along with this program.  If not, see [http://www.gnu.org/licenses/]. Contact me at thierry.maillard500n@orange.fr
