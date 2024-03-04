#####################################################################
# Document To PDF converter PDF to Document                         #
#                                                                   #
# Ce script convertit des feuilles Excel en documents PDF.          #
#                                                                   #
# Licence: MIT                                                      #
#                                                                   #
# Auteur: Florian Vaissiere                                         #
# GitHub: https://github.com/Askanat                                #
# Gitea: https://gitea.askanat.com                                  #
# LinkedIn: www.linkedin.com/in/florian-vaissiere-2bab64122         #
#####################################################################

import win32com.client
import os
import sys
import comtypes.client
from pdf2docx import Converter

def convert_sheet_to_pdf(sheet, output_path):
    """Exporte une feuille Excel donnée en PDF au chemin spécifié."""
    # Appliquer les paramètres de mise en page
    sheet.PageSetup.Orientation = win32com.client.constants.xlLandscape
    sheet.PageSetup.Zoom = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.FitToPagesTall = False

    # Exporter directement la feuille en PDF
    sheet.ExportAsFixedFormat(Type=win32com.client.constants.xlTypePDF, Filename=output_path)

def convert_excel_to_pdf(source_path, output_folder):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Rend l'application Excel invisible

    try:
        workbook = excel.Workbooks.Open(source_path)
        
        # Itérer sur chaque feuille et l'exporter en PDF
        for sheet in workbook.Worksheets:
            output_path = os.path.join(output_folder, f"{sheet.Name}.pdf")
            convert_sheet_to_pdf(sheet, output_path)

    except Exception as e:
        print(f"Erreur lors de la conversion : {e}")
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()

def convert_word_to_pdf(source_path, output_path):
    """Convertit un document Word en PDF."""
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(source_path)
    doc.SaveAs(output_path, FileFormat=17)  # 17 représente le format PDF dans Word
    doc.Close()
    word.Quit()

def convert_pdf_to_word(source_path, output_path):
    """Convertit un PDF en document Word."""
    cv = Converter(source_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def auto_convert_file(source_path, output_folder):
    """Détermine le type de fichier source et effectue la conversion appropriée."""
    _, ext = os.path.splitext(source_path)
    output_path = os.path.join(output_folder, os.path.splitext(os.path.basename(source_path))[0])

    if ext.lower() in ['.xls', '.xlsx']:
        convert_excel_to_pdf(source_path, output_folder)
    elif ext.lower() in ['.doc', '.docx']:
        convert_word_to_pdf(source_path, f"{output_path}.pdf")
    elif ext.lower() == '.pdf':
        convert_pdf_to_word(source_path, f"{output_path}.docx")
    else:
        print("Format de fichier non pris en charge.")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: script.py <source_path> <output_folder>")
        sys.exit(1)
    
    source_path = sys.argv[1]
    output_folder = sys.argv[2]

    auto_convert_file(source_path, output_folder)