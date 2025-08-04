import os
import re
import csv
import PyPDF2
import pdfplumber
import unicodedata
import shutil

# ====================================================================================================
# PDF Order Processing Logic (Core Module)
# This module contains the core functions for processing PDF order forms.
# It's designed to be imported and used by other applications (e.g., a Streamlit frontend).
# ====================================================================================================

# X-coordinates for table column splits (in points from left edge of the page).
# EDEKA-spezifische vertikale Linien
EDEKA_VERTICAL_LINES = [11.34, 99.21, 192.75, 272.12, 334.58, 411.02, 623.61, 722.82]

# Bounding Box for the area where tables are expected on each page [x0, y0, x1, y1] in points.
# x0, y0 = top-left corner; x1, y1 = bottom-right corner.
# EDEKA-spezifische BBoxes
EDEKA_BBOX_PAGE1 = [8.50, 314.63, 737.00, 524.30]
EDEKA_BBOX_OTHER_PAGES = [8.50, 99.21, 722.82, 538.57]

# Dohle HIT (AEZ)-spezifische vertikale Linien (umgerechnet von mm in Punkte)
DOHLE_VERTICAL_LINES = [28.35, 60.94, 216.74, 256.44, 279.11, 328.62, 390.87, 432.34, 462.98, 493.63, 527.29, 566.93]

# Dohle HIT (AEZ)-spezifische BBox (umgerechnet von mm in Punkte)
# WICHTIG: Diese BBox ist für einseitige Bestellungen. Falls mehrseitige AEZ-Bestellungen auftreten,
# muss eine DOHLE_BBOX_OTHER_PAGES definiert und die Logik in process_single_pdf angepasst werden.
DOHLE_BBOX = [28.35, 342.98, 566.93, 810.20]

# Mapping of a distinctive keyword from the delivery address to its corresponding market password.
# The keys in this dictionary must be in ALL CAPS and normalized according to the `clean_text` function
# (e.g., German umlauts ä, ö, ü replaced by AE, OE, UE; ß replaced by SS).
# Example: "An der Schäferwiese" -> "SCHAEFERWIESE", "Einsteinstraße" -> "EINSTEINSTRASSE"
MARKET_PASSWORDS = {
    "SCHAEFERWIESE": "Allee",         # Corresponds to "An der Schäferwiese"
    "THERESIENHOEHE": "Theresie",     # Corresponds to "Theresienhöhe"
    "EINSTEINSTRASSE": "Einstein",    # Corresponds to "Einsteinstraße" or "Einsteinstrasse"
    "UNTERHACHING": "Unterhaching", # Placeholder - replace with actual keywords for other markets
    "PULLACH": "Pullach",             # Placeholder - replace with actual keywords for other markets
    # NEUE PLATZHALTER FÜR DOHLE (AEZ)-MÄRKTE BASIEREND AUF DEM LETZTEN WORT IN "AEZ Haus XX NAME"
    "ISARTAL": "DohlePasswortIsartal",      # Beispiel für "AEZ Haus 80 Isartal"
    "MARTINSRIED": "DohlePasswortMartinsried", # Beispiel für "AEZ Haus 60 Martinsried"
    # Fügen Sie hier weitere Dohle-Märkte hinzu, falls nötig
}

# --- Helper Functions ---

def clean_text(text):
    """
    Cleans and normalizes text by:
    1. Replacing German umlauts (ä, ö, ü) with AE, OE, UE.
    2. Replacing 'ß' (sharp S) with 'ss'.
    3. Normalizing Unicode characters to their closest ASCII equivalent and removing non-ASCII characters.
    4. Removing extra whitespace and stripping leading/trailing spaces.

    Args:
        text (str): The input text string.

    Returns:
        str: The cleaned and normalized text.
    """
    if text is None:
        return ""

    # Explicit replacements for German special characters before general Unicode normalization.
    # This ensures consistent conversion (e.g., ä -> ae) as required for MARKET_PASSWORDS keys.
    text = text.replace('Ä', 'AE').replace('Ö', 'OE').replace('Ü', 'UE')
    text = text.replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue')
    text = text.replace('ẞ', 'SS').replace('ß', 'ss') # Handle both capital and small sharp S

    # Normalize Unicode characters (e.g., accented characters) to their base forms
    # and then encode to ASCII, ignoring any characters that cannot be represented.
    # This helps in handling a wide range of character encoding variations from PDFs.
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')

    # Replace multiple whitespace characters (including newlines) with a single space
    # and remove leading/trailing spaces.
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extract_info_from_text(text_content, is_dohle=False):
    """
    Extrahieren Bestelldatum, Bestellnummer, Lieferdatum, Markt-Kennwort und Marktname
    aus dem Roh-Text, der aus dem PDF-Header extrahiert wurde.

    Args:
        text_content (str): Der vollständige Textinhalt der PDF-Seite(n).
        is_dohle (bool): Gibt an, ob es sich um eine Dohle HIT (AEZ)-Bestellung handelt.

    Returns:
        tuple: (bestellnummer, bestelldatum, lieferdatum, markt_kennwort, markt_name)
    """
    bestellnummer = ""
    bestelldatum = ""
    lieferdatum = ""
    markt_kennwort = ""
    markt_name = "" # Nur für Dohle HIT relevant, aber immer zurückgeben

    if not is_dohle: # Logik für EDEKA-Bestellungen
        # Extraktion Bestelldatum (EDEKA)
        match = re.search(r"Bestelldatum:\s*(\d{2}\.\d{2}\.\d{4})", text_content)
        if match:
            bestelldatum = match.group(1)

        # Extraktion Liefertermin (EDEKA)
        match = re.search(r"Liefertermin:\s*(\d{2}\.\d{2}\.\d{4})", text_content)
        if match:
            lieferdatum = match.group(1)

        # Extraktion Bestellnummer (EDEKA)
        match = re.search(r"Bestellnummer:\s*(\d+)", text_content)
        if match:
            bestellnummer = match.group(1)
        # Zusätzlicher Regex für Bestellnummern wie "L 001/0002"
        if not bestellnummer: # Falls erste Regex nichts findet
            match = re.search(r"Bestell-Nr\.:\s*([A-Za-z0-9\s\/]+)", text_content)
            if match:
                bestellnummer = match.group(1).strip()

        # Extraktion Markt-Kennwort (EDEKA)
        # Extract the delivery address block for market identification.
        delivery_address_block = ""
        match_address = re.search(
            r'LIEFERANSCHRIFT\s*([\s\S]*?)(?=\n(?:GLN:|Empf\.:|Ihr Ansprechpartner\/in|RECHNUNGSEMPFÄNGER|$))',
            text_content,
            re.IGNORECASE
        )
        if match_address:
            delivery_address_block_raw = match_address.group(1)
            # Clean and uppercase the extracted address block for robust keyword matching
            delivery_address_block = clean_text(delivery_address_block_raw).upper()

            # Iterate through defined market passwords to find a matching keyword in the address block
            for key_fragment, password_value in MARKET_PASSWORDS.items():
                if key_fragment in delivery_address_block:
                    markt_kennwort = password_value
                    break

    else: # Logik für AEZ (Dohle HIT)-Bestellungen
        # Extraktion Bestellnummer (Dohle HIT)
        match = re.search(r"Bestellung Nr\.?\s*(\d+)", text_content)
        if match:
            bestellnummer = match.group(1)

        # Extraktion Bestelldatum (Dohle HIT)
        match = re.search(r"Datum:\s*(\d{2}\.\d{2}\.\d{4})", text_content)
        if match:
            bestelldatum = match.group(1)

        # Extraktion Liefertermin (Dohle HIT)
        match = re.search(r"Liefertermin:\s*(\d{2}\.\d{2}\.\d{4})", text_content)
        if match:
            lieferdatum = match.group(1)

        # Extraktion Marktname (Dohle HIT) - wird nicht in CSV ausgegeben, nur zur Erkennung/Info
        # Suchen nach "AEZ Haus XX NAME"
        # NEU: Das Regex wurde präzisiert, um nur den tatsächlichen Marktnamen zu erfassen und
        # angehängte Wörter wie "Seite" zu vermeiden.
        match = re.search(r"AEZ Haus \d+\s*([A-Za-zäöüÄÖÜ\s-]+?)(?=\s+GLN:|\n|$)", text_content)
        if match:
            markt_name_raw = match.group(1).strip() # Nur den Namensteil nehmen, z.B. "80 Isartal"
            # Versuche, die Nummer und den Namen zu trennen
            num_match = re.match(r"(\d+)\s*(.*)", markt_name_raw)
            if num_match:
                market_num = num_match.group(1)
                market_actual_name = num_match.group(2).strip()
                markt_name = f"AEZ Haus {market_actual_name}" # Setze den Namen ohne Nummer zurück, z.B. "AEZ Haus Isartal"
            else:
                markt_name = f"AEZ Haus {markt_name_raw}" # Fallback, falls keine Nummer gefunden wird

            # NEU: Extrahiere Schlüsselwort aus market_actual_name für die Marktpasswort-Suche
            # Das Schlüsselwort ist jetzt der letzte Teil des bereinigten Markennamens (z.B. "ISARTAL")
            cleaned_market_name_for_keyword = clean_text(markt_name_raw).upper() # Bereinige nur den erfassten Namensteil
            market_identifier_match = re.search(r'(\S+)$', cleaned_market_name_for_keyword)
            if market_identifier_match:
                extracted_keyword = market_identifier_match.group(1) # Z.B. "ISARTAL"
                if extracted_keyword in MARKET_PASSWORDS:
                    markt_kennwort = MARKET_PASSWORDS[extracted_keyword]

        # Wenn kein Markt-Kennwort gefunden wurde (oder keine Logik definiert), bleibt es leer
        if not markt_kennwort:
            markt_kennwort = ""

    return bestellnummer, bestelldatum, lieferdatum, markt_kennwort, markt_name

def process_single_pdf(pdf_input_path, temp_csv_output_dir, final_csv_download_dir):
    """
    Processes a single PDF file:
    1. Extracts full text for header information.
    2. Extracts table data using pdfplumber based on defined bounding boxes and vertical lines.
    3. Filters and processes article data.
    4. Generates a formatted CSV file.
    5. Copies the generated CSV to a final directory and deletes the temporary file.

    Args:
        pdf_input_path (str): The full path to the input PDF file.
        temp_csv_output_dir (str): Path to the temporary CSV output directory.
        final_csv_download_dir (str): Path to the final CSV download directory.

    Returns:
        str or None: Path to the generated CSV file on success, None on failure.
    """
    print(f"DEBUG: Start processing {os.path.basename(pdf_input_path)}")
    full_text = ""
    table_data_raw = []

    # Attempt to extract full text from the PDF using PyPDF2
    try:
        print(f"DEBUG: Attempting to open PDF with PyPDF2: {pdf_input_path}")
        with open(pdf_input_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            if reader.is_encrypted:
                print(f"DEBUG: PDF is encrypted.")
                print(f"WARNUNG: PDF ist verschlüsselt und kann nicht gelesen werden: {os.path.basename(pdf_input_path)}")
                return None
            print(f"DEBUG: PyPDF2 opened PDF. Number of pages: {len(reader.pages)}")
            for page_num in range(len(reader.pages)):
                full_text += reader.pages[page_num].extract_text() or ""
            print(f"DEBUG: PyPDF2 extracted full text.")
    except Exception as e:
        print(f"DEBUG: Caught exception during PyPDF2 processing: {e}")
        print(f"FEHLER beim Lesen des PDF-Textes von {os.path.basename(pdf_input_path)}: {e}")
        return None

    # Detect if it's a DOHLEHIT (AEZ) file based on filename
    file_basename_upper = os.path.basename(pdf_input_path).upper()
    is_dohle = "DOHLEHIT" in file_basename_upper or "AEZ" in file_basename_upper

    # Extract header information (date, order number, delivery date, market password, market name)
    bestellnummer, bestelldatum, lieferdatum, markt_kennwort, markt_name = extract_info_from_text(full_text, is_dohle=is_dohle)
    print(f"DEBUG: Header Info extracted - Bestellnummer: {bestellnummer}, Bestelldatum: {bestelldatum}, Lieferdatum: {lieferdatum}, Markt-Kennwort: '{markt_kennwort}', Marktname: '{markt_name}'")

    # Only EDEKA files require a market password to proceed. Dohle files do not (even if a password is found, it's optional for process continuation).
    if not is_dohle and not markt_kennwort:
        print(f"DEBUG: EDEKA-File without Market Password. Skipping.")
        print(f"WARNUNG: Markt-Kennwort konnte für EDEKA-Bestellung '{os.path.basename(pdf_input_path)}' nicht gefunden werden. Überspringe Datei.")
        return None

    # Determine BBOX and vertical lines based on file type
    bbox_to_use_page1 = None
    bbox_to_use_other_pages = None
    vertical_lines_to_use = None

    if is_dohle:
        bbox_to_use_page1 = DOHLE_BBOX
        bbox_to_use_other_pages = DOHLE_BBOX # Aktuell gleiche BBox für alle Seiten bei Dohle HIT
                                             # Bei mehrseitigen DohleHITs muss DOHLE_BBOX_OTHER_PAGES hierher
        vertical_lines_to_use = DOHLE_VERTICAL_LINES
        print(f"DEBUG: Using DOHLE BBOXes and Vertical Lines.")
    else: # Wenn nicht Dohle, dann EDEKA
        bbox_to_use_page1 = EDEKA_BBOX_PAGE1
        bbox_to_use_other_pages = EDEKA_BBOX_OTHER_PAGES
        vertical_lines_to_use = EDEKA_VERTICAL_LINES
        print(f"DEBUG: Using EDEKA BBOXes and Vertical Lines.")

    # Extract table data using pdfplumber for both EDEKA and Dohle files
    print(f"DEBUG: Attempting to open PDF with pdfplumber: {pdf_input_path}")
    try:
        with pdfplumber.open(pdf_input_path) as pdf:
            print(f"DEBUG: pdfplumber opened PDF. Number of pages: {len(pdf.pages)}")
            for page in pdf.pages:
                current_bbox = bbox_to_use_page1 if page.page_number == 1 else bbox_to_use_other_pages

                print(f"DEBUG: Processing page {page.page_number} with BBOX {current_bbox} and VLines {vertical_lines_to_use}.")
                # Sicherheitsprüfung für definierte BBoxes/Linien
                if current_bbox is None or vertical_lines_to_use is None:
                    print(f"DEBUG: Bounding Box or vertical lines are None. Cannot extract tables.")
                    print(f"FEHLER: Bounding Box oder vertikale Linien nicht für den Dateityp '{os.path.basename(pdf_input_path)}' definiert. Überspringe Datei.")
                    return None

                tables = page.crop((current_bbox[0], current_bbox[1], current_bbox[2], current_bbox[3])).extract_tables({
                    "vertical_strategy": "explicit", # Use defined vertical lines for column separation
                    "horizontal_strategy": "text",   # Detect horizontal lines based on text position
                    "explicit_vertical_lines": vertical_lines_to_use,
                    "min_words_horizontal": 1        # Minimum words to consider a horizontal line
                })
                print(f"DEBUG: Extracted {len(tables)} tables from page {page.page_number}.")
                for table in tables:
                    for row in table:
                        # Clean each cell's text upon extraction
                        table_data_raw.append([clean_text(cell) for cell in row])
            print(f"DEBUG: pdfplumber table extraction completed. Raw table data rows: {len(table_data_raw)}")

    except Exception as e:
        print(f"DEBUG: Caught exception during pdfplumber processing: {e}")
        print(f"FEHLER beim Extrahieren der Tabellen von {os.path.basename(pdf_input_path)}: {e}")
        return None

    if not table_data_raw: # Wenn keine Tabellendaten gefunden wurden
        print(f"DEBUG: No raw table data found after pdfplumber extraction.")
        print(f"WARNUNG: Keine passende Artikeltabelle in {os.path.basename(pdf_input_path)} gefunden. Prüfen Sie 'vertical_lines' oder 'BBOX_PAGE1'/'BBOX_OTHER_PAGES' und das PDF-Layout.")
        return None

    # Process and filter extracted article data (relevant for EDEKA and Dohle HIT)
    processed_article_data = []

    # EDEKA-spezifische Logik für den Start des Datenbereichs
    if not is_dohle:
        in_data_section = False
        print(f"DEBUG: EDEKA-specific data section handling enabled.")
    else:
        print(f"DEBUG: Dohle-specific data section handling (no 'in_data_section' flag).")


    print(f"DEBUG: Starting article data processing loop.")

    for i, row in enumerate(table_data_raw):
        print(f"DEBUG: Processing row {i}: {row}") # Print the raw row after clean_text
        row_joined_upper = " ".join(row).upper()

        # Skip empty rows
        if not any(cell.strip() for cell in row):
            print(f"DEBUG: Row {i} skipped: empty.")
            continue

        if not is_dohle: # EDEKA-specific logic for detecting data section start
            # Detect the start of the data section (e.g., a row with all underscores)
            if all(cell.strip().startswith('_') for cell in row if cell.strip()):
                in_data_section = True
                print(f"DEBUG: Row {i}: EDEKA data section STARTED.")
                continue

            # Only process rows if we are past the initial headers and in the data section for EDEKA
            if not in_data_section:
                print(f"DEBUG: Row {i} skipped: EDEKA, not in data section yet.")
                continue
            print(f"DEBUG: Row {i}: EDEKA, IN data section.")

        # Skip summary rows or irrelevant headers (applies to both EDEKA and Dohle)
        if "PLAN MHD" in row_joined_upper or "SUMME" in row_joined_upper or "GESAMT" in row_joined_upper:
            print(f"DEBUG: Row {i} skipped: Contains summary keyword.")
            continue

        # Column indices for article number and order quantity differ between EDEKA and DOHLE
        if is_dohle:
            # KORREKTUR: Lief.-Artikelnr. ist in Spalte 4, nicht 5. Menge ist in Spalte 7.
            lief_artnr_str = row[4] if len(row) > 4 else ""
            bestellmenge_str = row[7] if len(row) > 7 else ""
            print(f"DEBUG: Row {i} (Dohle): Lief.ArtNr='{lief_artnr_str}', Menge='{bestellmenge_str}'")
        else: # EDEKA Spalten (basierend auf der Reihenfolge der EDEKA_VERTICAL_LINES)
            lief_artnr_str = row[0] if len(row) > 0 else ""
            bestellmenge_str = row[1] if len(row) > 1 else ""
            print(f"DEBUG: Row {i} (EDEKA): Lief.ArtNr='{lief_artnr_str}', Menge='{bestellmenge_str}'")

        current_lief_artnr = ""
        # Validate that the article number is numeric
        if lief_artnr_str.strip().isdigit():
            current_lief_artnr = lief_artnr_str.strip()
            print(f"DEBUG: Row {i}: Article number '{current_lief_artnr}' is numeric.")
        else:
            print(f"DEBUG: Row {i} skipped: Article number '{lief_artnr_str}' is NOT numeric.")
            continue

        try:
            # Convert order quantity to float, handling comma as decimal separator
            bestellmenge = float(bestellmenge_str.replace(',', '.').strip())
            print(f"DEBUG: Row {i}: Quantity '{bestellmenge_str}' converted to {bestellmenge}.")

            # Add valid article data to the list
            # KORREKTUR: lieferdatum ist für die Artikelerfassung nicht zwingend notwendig.
            # Wenn es fehlt, wird es im CSV-Header einfach leer gelassen.
            if bestelldatum and bestellnummer and current_lief_artnr and bestellmenge > 0:
                    processed_article_data.append({
                        "artikelnummer": current_lief_artnr,
                        "bestellmenge": bestellmenge
                    })
                    print(f"DEBUG: Row {i}: Article added - {current_lief_artnr}, {bestellmenge}.")
            else:
                print(f"DEBUG: Row {i} skipped: Missing header info or quantity <= 0: "
                      f"BD={bool(bestelldatum)}, BN={bool(bestellnummer)}, LD={bool(lieferdatum)}, "
                      f"AL={bool(current_lief_artnr)}, Qty>0={bestellmenge > 0}.")
        except (ValueError, IndexError) as ve:
            print(f"DEBUG: Row {i} skipped: Value/Index error during quantity conversion ('{bestellmenge_str}') - {ve}.")
            continue

    if not processed_article_data:
        print(f"DEBUG: No valid article data found after filtering.")
        print(f"WARNUNG: Nach Filterung keine gültigen Artikeldaten für {os.path.basename(pdf_input_path)} gefunden.")
        return None
    print(f"DEBUG: Processed {len(processed_article_data)} valid articles.")

    # --- CSV File Generation ---
    output_filename = os.path.splitext(os.path.basename(pdf_input_path))[0] + ".csv"
    temp_csv_path = os.path.join(temp_csv_output_dir, output_filename)
    final_csv_path = os.path.join(final_csv_download_dir, output_filename)

    print(f"DEBUG: Preparing to write CSV to {temp_csv_path}")
    with open(temp_csv_path, 'w', newline='', encoding='ISO-8859-1') as csvfile:
        csv_writer = csv.writer(csvfile, delimiter=';')

        # Write header row (e.g., with market password in column 15 for EDEKA/empty for Dohle)
        header_row_csv = [''] * 15
        # Platzierung des Markt-Kennworts (für EDEKA) oder leer (für Dohle HIT) in Spalte O1 (Index 14)
        header_row_csv[14] = markt_kennwort
        csv_writer.writerow(header_row_csv)
        print(f"DEBUG: CSV header row written.")

        # Write empty row (as per desired CSV format)
        csv_writer.writerow([''] * 15)

        # Write Bestelldatum in column 3
        row_c3 = [''] * 15
        row_c3[2] = bestelldatum
        csv_writer.writerow(row_c3)

        # Write Bestellnummer in column 3
        row_c4 = [''] * 15
        row_c4[2] = bestellnummer
        csv_writer.writerow(row_c4)

        # Write Lieferdatum in column 3
        row_c5 = [''] * 15
        row_c5[2] = lieferdatum
        csv_writer.writerow(row_c5)
        print(f"DEBUG: CSV header data written.")

        # Write article data rows (only if processed_article_data is not empty)
        for item in processed_article_data:
            row_csv = [''] * 15
            row_csv[1] = item["artikelnummer"]
            row_csv[4] = str(item["bestellmenge"]).replace('.', ',') # Format quantity with comma for CSV
            csv_writer.writerow(row_csv)
        print(f"DEBUG: CSV article data written.")

    # Copy the temporary CSV to the final download directory
    try:
        print(f"DEBUG: Attempting to copy '{temp_csv_path}' to '{final_csv_path}'")
        shutil.copy(temp_csv_path, final_csv_path)
        print(f"CSV-Datei erfolgreich erstellt und kopiert nach: {final_csv_path}")

        # Delete the temporary CSV file after successful copy
        print(f"DEBUG: Attempting to delete temporary file '{temp_csv_path}'")
        os.remove(temp_csv_path)
        print(f"Temporäre Datei '{os.path.basename(temp_csv_path)}' gelöscht.")

        return final_csv_path # Return the path to the successfully created CSV

    except Exception as e:
        print(f"DEBUG: Caught exception during CSV copy/delete: {e}")
        print(f"FEHLER beim Kopieren oder Löschen der CSV-Datei: {e}")
        return None

# No main() function here anymore, as this file is imported as a module.