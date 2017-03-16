# coding: utf-8
from docx import Document
from docx.shared import Inches
from docx.enum.style import WD_STYLE_TYPE
from doc_generator import (setDocumentStyle, 
	generateDocumentTitle, generateDocumentSpecialNotes)

doc = Document()

doc = setDocumentStyle(doc)
doc = generateDocumentTitle(doc)
doc = generateDocumentSpecialNotes(doc)

doc.save("demo.docx")
