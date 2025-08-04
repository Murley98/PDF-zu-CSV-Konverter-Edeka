import streamlit as st
import zipfile
import os
import shutil
import tempfile # For creating temporary directories
from pdf_processing_logic import process_single_pdf # Import the processing function

# ====================================================================================================
# Streamlit Web Application for PDF Order Conversion
# This script provides a simple drag-and-drop interface to convert PDF order forms to CSV.
# It utilizes the core processing logic from 'pdf_processing_logic.py'.
# ====================================================================================================

st.set_page_config(layout="wide") # Use wide layout for a larger drag-and-drop area

st.title("📄 Edeka Bestungen Konverter")

st.markdown("""
    Dieses Tool konvertiert die Edeka bestell PDF´s automatisch in die Passenden CSV Dateien für CSB.

    **So funktioniert's:**
    1.  Ziehen Sie Ihre PDF-Bestellungsdateien in das Drag & Drop-Feld unten.
    2.  Klicken Sie auf den "Dateien konvertieren"-Button.
    3.  Laden Sie die generierten CSV-Dateien direkt hier herunter.
""")

# Drag & Drop / File Uploader
uploaded_files = st.file_uploader(
    "**PDF-Dateien hierher ziehen oder zum Hochladen klicken**", 
    type="pdf", 
    accept_multiple_files=True,
    key="pdf_uploader_widget" # Unique key for the widget
)

# Convert Button
if st.button("**Dateien konvertieren**", key="convert_button"):
    if not uploaded_files:
        st.warning("Bitte laden Sie zuerst mindestens eine PDF-Datei hoch.")
    else:
        st.info("Verarbeitung gestartet... Bitte warten Sie. Dies kann je nach Dateigröße und Anzahl etwas dauern.")

        # Create a unique temporary root directory for this session's processing.
        # This ensures that files from different user sessions (or runs) don't interfere
        # and allows for easy cleanup of all related temporary files.
        with tempfile.TemporaryDirectory() as temp_root_dir:
            # Define subdirectories within the temporary root for input, temporary CSVs, and final CSVs.
            temp_pdf_input_dir = os.path.join(temp_root_dir, "input_pdfs")
            temp_csv_output_dir = os.path.join(temp_root_dir, "temp_csvs")
            final_csv_download_dir = os.path.join(temp_root_dir, "final_csvs")

            # Ensure these temporary directories exist
            os.makedirs(temp_pdf_input_dir, exist_ok=True)
            os.makedirs(temp_csv_output_dir, exist_ok=True)
            os.makedirs(final_csv_download_dir, exist_ok=True)

            processed_csv_paths = [] # To store paths of successfully generated CSVs

            for uploaded_file in uploaded_files:
                # Display current processing status
                st.text(f"Verarbeite: {uploaded_file.name}...")

                # Save the Streamlit UploadedFile object to a physical file in the temporary input directory.
                # The core processing logic needs a file path, not an UploadedFile object directly.
                input_pdf_path = os.path.join(temp_pdf_input_dir, uploaded_file.name)
                with open(input_pdf_path, "wb") as f:
                    f.write(uploaded_file.getbuffer()) # .getbuffer() gets the bytes from the uploaded file

                try:
                    # Call the core processing logic for the saved PDF.
                    # Pass the dynamic temporary directories for output.
                    final_csv_path = process_single_pdf(
                        input_pdf_path, 
                        temp_csv_output_dir, 
                        final_csv_download_dir
                    )

                    if final_csv_path:
                        processed_csv_paths.append(final_csv_path)
                        st.success(f"'{uploaded_file.name}' erfolgreich konvertiert.")
                    else:
                        st.warning(f"Konnte '{uploaded_file.name}' nicht verarbeiten. Bitte prüfen Sie das Dateiformat.")
                except Exception as e:
                    # Catch any unexpected errors during processing and display them.
                    st.error(f"Ein unerwarteter Fehler ist bei der Verarbeitung von '{uploaded_file.name}' aufgetreten: {e}")

        # --- Prepare and Display ZIP Download for Processed CSVs ---
        if processed_csv_paths:
            st.subheader("Ihre konvertierten CSV-Dateien als ZIP-Datei:")

            # Create a temporary ZIP file path within the main temp directory
            zip_file_name = "konvertierte_bestellungen.zip"
            # temp_root_dir kommt vom `with tempfile.TemporaryDirectory() as temp_root_dir:` Block
            zip_file_path = os.path.join(temp_root_dir, zip_file_name) 

            try:
                with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for csv_path in processed_csv_paths:
                        # Add each CSV file to the ZIP archive
                        # Use os.path.basename for the name inside the zip
                        zipf.write(csv_path, os.path.basename(csv_path))

                # Read the created ZIP file to offer it for download
                with open(zip_file_path, "rb") as f:
                    st.download_button(
                        label="📦 Alle CSVs als ZIP herunterladen",
                        data=f.read(),
                        file_name=zip_file_name,
                        mime="application/zip",
                        key="download_all_csvs_zip"
                    )
                st.success(f"Alle {len(processed_csv_paths)} CSV-Dateien wurden in '{zip_file_name}' verpackt und stehen zum Download bereit.")

            except Exception as e:
                st.error(f"FEHLER beim Erstellen der ZIP-Datei: {e}")
                st.error("Es konnten keine Dateien erfolgreich konvertiert werden oder die ZIP-Erstellung schlug fehl. Bitte überprüfen Sie die hochgeladenen PDFs und die Konfiguration.")
        else:
            st.warning("Es wurden keine CSV-Dateien generiert oder gefunden.") # Keep this warning if no files were processed successfully

        st.info("Alle temporären Dateien wurden bereinigt.")
        # The 'temp_root_dir' and all its contents are automatically deleted when exiting the 'with' block.

# Footer for additional information
st.markdown("---")
st.markdown("Ein Tool bereitgestellt von Simon Murr.")