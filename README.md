# Fin_Acronym_Dictionary
Excel/VBA spreadsheet that creates a list of acronyms from pasted text

PURPOSE: This Excel application can take any text, parse it for Finance acronyms, and return a formatted glossary of the acronyms used within the text.

DESCRIPTION OF FILES:

Dictionary_App.xlsm - contains complete program already populated with 1000+ acronyms. Developed on Excel for Mac &
therefore lacks some ActiveX features that would improve the UI. Just have to paste in the text and it will create
a glossary of the acronyms used within the text

VBA_App.bas - extract of the VB code used to create the program

Scrape_Terms.py - code used to scrape data from Investopedia and populate the acronym dictionary within the app
Requires installation of Pandas

