import win32com.client
import os
import sys

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

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: convert_excel_to_pdf.py <source_path> <output_folder>")
        sys.exit(1)
    
    source_path = sys.argv[1]
    output_folder = sys.argv[2]
    convert_excel_to_pdf(source_path, output_folder)
