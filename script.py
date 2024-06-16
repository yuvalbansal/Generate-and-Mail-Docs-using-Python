import re
from docx import Document
from docx2pdf import convert

def replace_template_variables(docx_file, replacements):
    # Load the document
    doc = Document(docx_file)

    # Iterate through paragraphs and replace variables
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if f'{{{{{key}}}}}' in para.text:
                para.text = para.text.replace(f'{{{{{key}}}}}', value)

    # Save the modified document
    modified_docx_file = 'modified_template.docx'
    doc.save(modified_docx_file)
    return modified_docx_file

def generate_pdf_from_template(docx_file, replacements, output_pdf):
    # Replace variables in the template
    modified_docx_file = replace_template_variables(docx_file, replacements)

    # Convert the modified document to PDF
    convert(modified_docx_file, output_pdf)

# Example usage
docx_file = 'Letter Template.docx'
replacements = {
    'Abstract Number': 'Number Value',
    'Primary Author Name': 'Name Value',
    'Abstract Title': 'Title Value'
}
output_pdf = 'output.pdf'

generate_pdf_from_template(docx_file, replacements, output_pdf)
print(f"PDF generated: {output_pdf}")