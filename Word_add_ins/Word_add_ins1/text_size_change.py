import os
from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.shared import Pt


def txt_size_change(filename):
    doc=Document(filename)#Document object

   
    for para in doc.paragraphs:
        for runli in para.runs:
            if (int(runli.font.size.pt))!=12:
                runli.font.size =Pt(12)

    doc.save("app_output\\txt_change_output.docx")
    os.system("start app_output\\txt_change_output.docx")
    


txt_size_change("app_input\\text_size.docx")
