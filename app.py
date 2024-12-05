import os
from pathlib import Path
import streamlit as st
import pypandoc

# Verificar si Pandoc está instalado y descargarlo si no lo está
if not pypandoc.get_pandoc_path():
    pypandoc.download_pandoc()

def convert_docx_to_pdf(docx_file, pdf_path):
    try:
        # Convertir el archivo .docx a .pdf usando pypandoc
        pypandoc.convert_file(docx_file, 'pdf', outputfile=pdf_path)
    except Exception as e:
        st.error(f"Error al convertir el archivo {docx_file} a PDF: {e}")

def main():
    st.markdown("""
        <h1 style="text-align: center;">Conversor de Word a PDF</h1>
        <h3 style="text-align: center;">Creado por Juancito Pena</h3>
        <p style="text-align: center;">Selecciona los archivos Word para convertirlos a PDF.</p>
    """, unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Elige archivos Word (.docx)", type=["docx"], accept_multiple_files=True)
    
    if uploaded_files:
        st.write(f"Archivos seleccionados: {len(uploaded_files)} documentos.")
    
    output_folder = st.text_input("Escribe la ruta de la carpeta de salida (opcional):", "")
    
    if st.button("Convertir a PDF"):
        if not uploaded_files:
            st.error("Por favor, selecciona al menos un archivo Word.")
            return
        
        output_folder = output_folder or os.getcwd()
        os.makedirs(output_folder, exist_ok=True)
        
        with st.spinner("Convirtiendo documentos..."):
            for uploaded_file in uploaded_files:
                temp_file_path = os.path.join(output_folder, uploaded_file.name)
                pdf_path = os.path.join(output_folder, f"{Path(uploaded_file.name).stem}.pdf")
                
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Convertir .docx a PDF
                convert_docx_to_pdf(temp_file_path, pdf_path)
                
                os.remove(temp_file_path)
            
            st.success("¡Conversión completada!")
            st.write(f"Los archivos PDF se han guardado en: `{output_folder}`")

if __name__ == "__main__":
    main()
