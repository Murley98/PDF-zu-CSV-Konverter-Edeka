import streamlit as st
import os
import shutil
import tempfile
from pdf_processing_logic import process_single_pdf # Import the processing function

st.set_page_config(layout="wide") # Use a wide layout for a larger drag-and-drop area

st.title("EDEKA Bestellungen Konverter")

st.markdown("""
### Dieses Tool konvertiert die EDEKA Bestell-PDF's automatisch in die Passenden CSV Dateien für CSV.
""")

st.markdown("""
**So funktioniert's:**
1. Ziehen Sie Ihre PDF-Bestellungsdateien in das Drag & Drop-Feld unten.
2. Klicken Sie auf den "Dateien konvertieren"-Button.
3. Laden Sie die generierten CSV-Dateien direkt hier herunter.
""")

st.markdown("---")

st.markdown("### Drag & Drop / File Uploader")
uploaded_files = st.file_uploader(
    "PDF-Dateien hierher ziehen oder zum Hochladen klicken*",
    type="pdf",
    accept_multiple_files=True,
    key="pdf_uploader_widget" # Unique key for the widget
)

# Liste zum Speichern aller konvertierten CSV-Daten
converted_csv_data = []

if st.button("Dateien konvertieren"):
    if uploaded_files:
        temp_dir = tempfile.mkdtemp()
        for uploaded_file in uploaded_files:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Verarbeiten der einzelnen PDF-Datei
            try:
                # Annahme: process_single_pdf gibt den CSV-Inhalt als String zurück
                csv_content = process_single_pdf(file_path)

                # Dateiname für CSV anpassen
                csv_filename = os.path.splitext(uploaded_file.name)[0] + ".csv"
                converted_csv_data.append((csv_filename, csv_content))

            except Exception as e:
                st.error(f"Fehler bei der Verarbeitung von {uploaded_file.name}: {e}")

        shutil.rmtree(temp_dir)
        st.success("Konvertierung abgeschlossen!")

        # Zeige alle Download-Buttons AN, NACHDEM ALLE DATEIEN VERARBEITET WURDEN
        if converted_csv_data:
            st.markdown("### Konvertierte CSV-Dateien zum Download:")
            for filename, content in converted_csv_data:
                st.download_button(
                    label=f"Download {filename}",
                    data=content,
                    file_name=filename,
                    mime="text/csv",
                    key=f"download_button_{filename.replace('.', '_')}" # Einzigartiger Schlüssel für jeden Button
                )
        else:
            st.warning("Es wurden keine CSV-Dateien generiert.")

    else:
        st.warning("Bitte laden Sie zuerst mindestens eine PDF-Datei hoch.")

# Optional: Anzeige für erfolgreich heruntergeladene Dateien (falls gewünscht)
# Dies ist schwieriger zu implementieren, da Streamlit bei Klick neu lädt.
# Die beste Lösung ist, alle Buttons anzuzeigen, wie oben beschrieben.