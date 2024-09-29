import os
import PyPDF2
import openpyxl
from tkinter import Tk
from tkinter.filedialog import askdirectory

def count_pdf_pages(pdf_path):
    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        return len(reader.pages)

def get_pdf_files_in_folder(folder_path):
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    return pdf_files

def write_to_excel(pdf_info, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "PDF Info"
    sheet.append(["Ruta", "Nombre del archivo", "Cantidad de páginas"])

    for info in pdf_info:
        sheet.append([info['path'], info['name'], info['pages']])
    
    workbook.save(output_file)

def main():
    # Selección de carpeta
    root = Tk()
    root.withdraw()
    folder_path = askdirectory(title="Selecciona la carpeta raíz")

    if not folder_path:
        print("No se seleccionó ninguna carpeta.")
        return
    
    # Obtener todos los archivos PDF
    pdf_files = get_pdf_files_in_folder(folder_path)
    
    # Extraer la cantidad de páginas de cada PDF
    pdf_info = []
    for pdf_file in pdf_files:
        try:
            pages = count_pdf_pages(pdf_file)
            pdf_info.append({
                'path': pdf_file,
                'name': os.path.basename(pdf_file),
                'pages': pages
            })
        except Exception as e:
            print(f"Error al leer {pdf_file}: {e}")
    
    # Guardar en Excel
    output_file = os.path.join(folder_path, "pdf_info.xlsx")
    write_to_excel(pdf_info, output_file)
    
    print(f"Información exportada a {output_file}")

if __name__ == "__main__":
    main()
