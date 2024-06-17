import csv
from docxtpl import DocxTemplate

accept_doc = DocxTemplate("Accepted Letter Template.docx")
reject_doc = DocxTemplate("Rejected Letter Template.docx")

Number = 999
Author = 'Dummy'
Title = 'Sample'
Comments = 'None'
context = {'Number': Number, 'Author': Author, 'Title': Title, 'Comments': Comments}

accept_doc.render(context)
accept_doc.save(f"{Number}_Accepted.docx")