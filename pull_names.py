import requests
import json
import os
import urllib
import sys
import pprint

from docx import Document
from docx.shared import Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor, Pt

"""
NOTES:

Possible Repeat of First or Last names in the assigned table columns. 
Reason: Table formatting, though correct when looked at in MS Word. From PythonDocx parser,
the formatting throws it off if it's not the same as what we expect of the others.

- JH 2/21/17

"""

# Open Document
doc = Document(sys.argv[1])

# Grab all Tables, assuming additionals files will have directory table from TEMPLATE CREATION
dir_tables = doc.tables
names_list = []


for table in dir_tables:
	# Column Cells index argument can/should be changed if there are all repeats of first or last names - JH 3/7/2017
	for first, last in zip(table.column_cells(1), table.column_cells(2)):
		if "Was this person's info updated" in first.text or "Was this person's info updated" in last.text:
			pass
		else:
			whole_name = first.text + " " + last.text
			names_list.append(whole_name)

print names_list
print len(names_list)

f = open("names.txt", "w")
f.write("Number of Members: " + str(len(names_list)) + "\n\n")
for name in names_list:
	try:
		f.write(name + "\n")
	except UnicodeEncodeError:
		utf_name = name.encode('utf-8')
		f.write(utf_name + "\n")
		print name

f.close()
