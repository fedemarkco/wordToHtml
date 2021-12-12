from win32com import client

import os


inFileT = "inFileT.docx"

name, ext = os.path.splitext(inFileT)

inFile = os.path.abspath(inFileT)
outFile = os.path.abspath(name + ".html")
word = client.Dispatch("Word.Application")
doc = word.Documents.Open(inFile)
doc.SaveAs(outFile, FileFormat=10)
doc.Close()
word.Quit()

