from tempfile import TemporaryDirectory
from functions.createfolder import *
from win32com import client
from mailmerge import MailMerge
from datetime import date
import sys
import glob
from PyQt5 import QtGui
from PyQt5 import QtWidgets
from win32com import client
from PyQt5.QtWidgets import QInputDialog,QWidget
import functions.mail as mail


def acceptance_no_deficiencies(self, variation, project, path, date, item, item2):
    path2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


    # Define the temporary directory path
    temp_dir = os.path.join(path2, "Templates", "Word")  # Create Temp folder path relative to the script directory

    # Build the template file path dynamically using os.path.join
    template_filename = f"{variation}-{project}-{item}-{item2} Letter-Template.docx"
    template = os.path.join(temp_dir, template_filename)

    # Print the resulting template path for debugging (optional)
    print(f"Template file path: {template}")
    document = MailMerge(template)
    print(document.get_merge_fields())
    
    body_text = "All {0} deficiencies are now satisfied or none were found.".format(variation)
    document.merge(Date = date, Work_Order = self.planfile_entry.text(), Body = body_text )
    print("This is the path before the write: "+ path)
    file_path = path + "/{0}-{1}-{2}-{3}-{4} (Letter).docx".format(variation,date, project, item, item2)
    pdf_path = path + "/{0}-{1}-{2}-{3}-{4}(Letter).pdf".format(variation, date,project, item, item2)
    document.write(file_path)
    document.close()

    wdFormatPDF = 17

    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    mail.mail_to_sign(self,pdf_path,path)


def rejected_word(self, variation, project, path, date, item, item2):
    # Define the temporary directory path
    path2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    temp_dir = os.path.join(path2, "Templates", "Word")  # Create Temp folder path relative to the script directory

    # Build the template file path dynamically using os.path.join
    template_filename = f"{variation}-{project}-{item}-{item2} Letter-Template.docx"
    template = os.path.join(temp_dir, template_filename)

    # Print the resulting template path for debugging (optional)
    print(f"Template file path: {template}")
    document = MailMerge(template)
    print(document.get_merge_fields())
    
    body_text = "Defeciencies were found. Please see attached."
    document.merge(Date = date, Work_Order = self.planfile_entry.text(), Body = body_text )
    print("This is the path before the write: "+ path)
    file_path = path + "/{0}-{1}-{2}-{3}-{4} (Letter).docx".format(variation,date, project, item, item2)
    pdf_path = path + "/{0}-{1}-{2}-{3}-{4}(Letter).pdf".format(variation, date,project, item, item2)
    document.write(file_path)
    document.close()

    wdFormatPDF = 17

    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    mail.mail_to_sign(self,pdf_path,path)
