import os
import re
import PyPDF2
import docx
from openpyxl import Workbook

def extract_text_from_pdf(pdf_file):
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text 


def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return '\n'.join(full_text)

def extract_emails_and_phones(text):
    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    phones = re.findall(r'\b(?:\+?(\d{1,3}))?[-. (]*(\d{3})[-. )]*(\d{3})[-. ]*(\d{4})\b', text)
    return emails , phones

def process_cv_files(directory):
    wb = Workbook()
    ws = wb.active
    ws.append(["File Name", "Phone Numbers", "Text"])

    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(os.path.join(directory , filename))
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(os.path.join(directory, filename))
        else: 
            continue

        emails, phones = extract_emails_and_phones(text)
        ws.append([filename, ", ".join(emails), ", ".join([f'({p[0]}) {p[1]}-{p[2]}-{p[3]}' for p in phones]), text])

    wb.save("cv_data.xlsx")

if __name__ == "__main__":
    cv_directory = "./Sample2"
    process_cv_files(cv_directory)

    ## task one ost platform web development internship 