import csv
from docxtpl import DocxTemplate
from docx2pdf import convert

accept_doc = DocxTemplate("Accepted Letter Template.docx")
reject_doc = DocxTemplate("Rejected Letter Template.docx")

with open('Review Abstract Submission.csv', mode ='r')as file:
  csvFile = csv.reader(file)
  for lines in csvFile:
        Number = lines[0]
        Author = lines[3]
        Title = lines[1]
        Comments = lines[5]
        context = {'Number': Number, 'Author': Author, 'Title': Title, 'Comments': Comments}
        if lines[4] == "Accepted":
            accept_doc.render(context)
            accept_doc.save(f"Letters_Docx\{Number}_Accepted.docx")
            convert(f"Letters_Docx\{Number}_Accepted.docx", f"Letters_PDF\{Number}_Accepted.pdf")
        elif lines[4] == "Rejected":
            reject_doc.render(context)
            reject_doc.save(f"Letters_Docx\{Number}_Rejected.docx")
            convert(f"Letters_Docx\{Number}_Rejected.docx", f"Letters_PDF\{Number}_Rejected.pdf")
 