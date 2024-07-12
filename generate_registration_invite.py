import csv
from docxtpl import DocxTemplate
from docx2pdf import convert

invite_doc = DocxTemplate("Registration Invite Template.docx")

with open('Review Abstract Submission.csv', mode ='r')as file:
  csvFile = csv.reader(file)
  for lines in csvFile:
        Number = lines[0]
        Author = lines[3]
        Title = lines[1]
        context = {'Number': Number, 'Author': Author, 'Title': Title}
        if (Number.isdigit() == True) and (int(Number) > 251):
          print("Generating for", Number)
          if lines[4] == "Accepted":
              invite_doc.render(context)
              invite_doc.save(f"Invite_Docx\{Number}_Invite.docx")
              convert(f"Invite_Docx\{Number}_Invite.docx", f"Invite_PDF\{Number}_Invite.pdf")