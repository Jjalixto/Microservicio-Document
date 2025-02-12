
#aun falta acoplarlo al proyecto

import os
import win32com.client

def convertir_word_a_pdf(word_file):
    # Verifica si el archivo de Word existe
    if not os.path.exists(word_file):
        print(f"El archivo {word_file} no existe.")
        return
    
    # Inicializa la aplicación de Microsoft Word
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False  # Evita que se abra la interfaz gráfica

    # Abre el archivo de Word
    doc = word.Documents.Open(word_file)
    
    # Elimina todos los comentarios del documento
    for comentario in doc.Comments:
        comentario.Delete()
    
    # Convierte el archivo a PDF
    pdf_file = word_file.replace(".docx", ".pdf")
    doc.SaveAs(pdf_file, FileFormat=17)  # 17 es el código de PDF en Word
    
    # Cierra el documento y Word
    doc.Close()
    word.Quit()
    
    print(f"Archivo PDF guardado como: {pdf_file}")

# Ejemplo de uso
convertir_word_a_pdf(r'C:\Users\JoelJalixtoChavez\Desktop\modelo.docx')
