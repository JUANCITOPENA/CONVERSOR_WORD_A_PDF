import os
import tempfile
import streamlit as st
import pythoncom
from docx2pdf import convert

def convert_docx_to_pdf(docx_file, pdf_path):
    try:
        pythoncom.CoInitialize()
        try:
            convert(docx_file, pdf_path)
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        st.error(f"Error al convertir {docx_file}: {e}")

def main():
    st.markdown("""
        <h1 style="text-align: center;">Conversor de Word a PDF</h1>
        <h3 style="text-align: center;">Creado por Juancito Pena</h3>
        <p style="text-align: center;">Selecciona los archivos Word para convertirlos a PDF.</p>
    """, unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Elige archivos Word (.docx)", type=["docx"], accept_multiple_files=True)
    output_folder = st.text_input("Escribe la ruta de la carpeta de salida (opcional):", "").strip()
    
    if st.button("Convertir a PDF"):
        if not uploaded_files:
            st.error("Por favor, selecciona al menos un archivo Word.")
            return
        
        output_folder = output_folder or os.getcwd()
        if not os.path.exists(output_folder):
            st.error(f"La carpeta de salida especificada no existe: {output_folder}")
            return
        
        os.makedirs(output_folder, exist_ok=True)
        
        with st.spinner("Convirtiendo documentos..."):
            for uploaded_file in uploaded_files:
                try:
                    # Crear archivo temporal
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                        temp_file.write(uploaded_file.getbuffer())
                        temp_file_path = temp_file.name

                    # Ruta de salida para el PDF
                    pdf_path = os.path.join(output_folder, f"{os.path.splitext(uploaded_file.name)[0]}.pdf")
                    
                    # Convertir .docx a PDF
                    convert_docx_to_pdf(temp_file_path, pdf_path)
                    
                finally:
                    # Eliminar archivo temporal
                    if os.path.exists(temp_file_path):
                        os.remove(temp_file_path)
            
            st.success("¡Conversión completada!")
            st.write(f"Los archivos PDF se han guardado en: `{output_folder}`")
            for uploaded_file in uploaded_files:
                pdf_name = f"{os.path.splitext(uploaded_file.name)[0]}.pdf"
                pdf_path = os.path.join(output_folder, pdf_name)
                st.markdown(f"[Descargar {pdf_name}]({pdf_path})", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
