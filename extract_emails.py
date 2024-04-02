import os
import re
import fitz  # PyMuPDF
from docx import Document

def extract_emails_from_text(text):
    """Extracts email addresses from text."""
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    return re.findall(email_pattern, text)

def extract_emails_from_pdf(pdf_path):
    """Extracts emails from a PDF file."""
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return extract_emails_from_text(text)

def extract_emails_from_docx(docx_path):
    """Extracts emails from a DOCX file."""
    doc = Document(docx_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text
    return extract_emails_from_text(text)

def extract_emails_from_folder(folder_path):
    """Goes through all PDF and DOCX in a folder and extracts emails."""
    emails = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                emails.extend(extract_emails_from_pdf(os.path.join(root, file)))
            elif file.endswith(".docx"):
                emails.extend(extract_emails_from_docx(os.path.join(root, file)))
    return set(emails)  # Using set to remove duplicates

# Specify your folder path here
folder_path = 'EnterFilePathHere'
emails = extract_emails_from_folder(folder_path)
for email in emails:
    print(email)
