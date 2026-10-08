import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

doc = Document(r'c:\Users\user\OneDrive\Desktop\geranium project\NovaMart_extracted\Inputs\Broker_Email_Correspondence.docx')
for i, p in enumerate(doc.paragraphs):
    print(f"P{i}: {repr(p.text)}")
sys.stdout.flush()
