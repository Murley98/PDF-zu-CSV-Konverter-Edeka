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
# These are crucial for accurate table extraction and must be adjusted if the PDF layout changes.
vertical_lines = [11.34, 99.21, 192.75, 272.12, 334.58, 411.02, 623.61, 722.82] 

# Bounding Box for the area where tables are expected on each page [x0, y0, x1, y1] in points.
# x0, y0 = top-left corner; x1, y1 = bottom-right corner.
# Page 1 often has a different layout (e.g., header information) than subsequent pages.
BBOX_PAGE1 = [8.50, 314.63, 737.00, 524.30]
BBOX_OTHER_PAGES = [8.50, 99.21, 722.82, 538.57]

# Mapping of a distinctive keyword from the delivery address to its corresponding market password.
# The keys in this dictionary must be in ALL CAPS and normalized according to the `clean_text` function
# (e.g., German umlauts ä, ö, ü replaced by AE, OE, UE; ß replaced by SS).
# Example: "An der Schäferwiese" -> "SCHAEFERWIESE", "Einsteinstraße" -> "EINSTEINSTRASSE"
MARKET_PASSWORDS = {
    "SCHAEFERWIESE": "Allee",        # Corresponds to "An der Schäferwiese"
    "THERESIENHOEHE": "Theresie",    # Corresponds to "Theresienhöhe"
    "EINSTEINSTRASSE": "Einstein",  # Corresponds to "Einsteinstraße" or "Einsteinstrasse"
    "UNTERHACHING": "Unterhaching", # Placeholder - replace with actual keywords for other markets
    "PULLACH": "Pullach",           # Placeholder - replace with actual keywords for other markets
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

def extract_info_from_text(text):
    """
    Extracts Bestelldatum, Bestellnummer, Lieferdatum, and identifies the market password
    from the raw text extracted from the PDF header.

    Args:
        text (str): The full text content of the PDF page.

    Returns:
        tuple: (bestelldatum, bestellnummer, lieferdatum, markt_kennwort)
    """
    bestelldatum = ""
    bestellnummer = ""
    lieferdatum = ""
    markt = ""

    # Regex patterns to find order details. Uses [\s\S]*? to match any characters
    # (including newlines) non-greedily between the keyword and the value.
    match = re.search(r'Bestelldatum:[\s\S]*?(\d{2}\.\d{2}\.\d{4})', text)
    if match:
        bestelldatum = match.group(1)

    match = re.search(r'Bestellnummer:[\s\S]*?(\d+)', text)
    if match:
        bestellnummer = match.group(1)

    match = re.search(r'Lieferdatum/-uhrzeit:[\s\S]*?(\d{2}\.\d{2}\.\d{4})', text)
    if match:
        lieferdatum = match.group(1)

    # Extract the delivery address block for market identification.
    # This regex captures text between "LIEFERANSCHRIFT" and the next known header
    # (GLN:, Empf.:, Ihr Ansprechpartner/in, RECHNUNGSEMPFÄNGER) or end of string.
    # Uses IGNORECASE for case-insensitive matching.
    delivery_address_block = ""
    match_address = re.search(
        r'LIEFERANSCHRIFT\s*([\s\S]*?)(?=\n(?:GLN:|Empf\.:|Ihr Ansprechpartner\/in|RECHNUNGSEMPFÄNGER|$))', 
        text, 
        re.IGNORECASE
    )
    if match_address:
        delivery_address_block_raw = match_address.group(1)
        # Clean and uppercase the extracted address block for robust keyword matching
        delivery_address_block = clean_text(delivery_address_block_raw).upper() 

        # Iterate through defined market passwords to find a matching keyword in the address block
        for key_fragment, password_value in MARKET_PASSWORDS.items():
            if key_fragment in delivery_address_block:
                markt = password_value
                break

    return bestelldatum, bestellnummer, lieferdatum, markt

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
    print(f"Verarbeite Datei: {os.path.basename(pdf_input_path)}")
    full_text = ""
    table_data_raw = []

    # Attempt to extract full text from the PDF using PyPDF2
    try:
        with open(pdf_input_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            if reader.is_encrypted:
                print(f"WARNUNG: PDF ist verschlüsselt und kann nicht gelesen werden: {os.path.basename(pdf_input_path)}")
                return None
            for page_num in range(len(reader.pages)):
                full_text += reader.pages[page_num].extract_text() or ""
    except Exception as e:
        print(f"FEHLER beim Lesen des PDF-Textes von {os.path.basename(pdf_input_path)}: {e}")
        return None

    # Extract header information (date, order number, delivery date, market password)
    bestelldatum, bestellnummer, lieferdatum, markt_kennwort = extract_info_from_text(full_text)

    # Check if it's a HAMMERER file (based on filename) or if a market password was found for EDEKA
    is_hammerer = "HAMMERER" in os.path.basename(pdf_input_path).upper()
    if not is_hammerer and not markt_kennwort:
        print(f"WARNUNG: Markt-Kennwort konnte für EDEKA-Bestellung '{os.path.basename(pdf_input_path)}' nicht gefunden werden. Überspringe Datei.")
        return None

    # Extract table data using pdfplumber
    try:
        with pdfplumber.open(pdf_input_path) as pdf:
            for page in pdf.pages:
                # Use different bounding boxes for page 1 vs. other pages if layout differs
                bbox_to_use = BBOX_PAGE1 if page.page_number == 1 else BBOX_OTHER_PAGES
                tables = page.crop((bbox_to_use[0], bbox_to_use[1], bbox_to_use[2], bbox_to_use[3])).extract_tables({
                    "vertical_strategy": "explicit", # Use defined vertical lines for column separation
                    "horizontal_strategy": "text",   # Detect horizontal lines based on text position
                    "explicit_vertical_lines": vertical_lines,
                    "min_words_horizontal": 1        # Minimum words to consider a horizontal line
                })
                for table in tables:
                    for row in table:
                        # Clean each cell's text upon extraction
                        table_data_raw.append([clean_text(cell) for cell in row])

    except Exception as e:
        print(f"FEHLER beim Extrahieren der Tabellen von {os.path.basename(pdf_input_path)}: {e}")
        return None

    if not table_data_raw:
        print(f"WARNUNG: Keine passende Artikeltabelle in {os.path.basename(pdf_input_path)} gefunden. Prüfen Sie 'vertical_lines' oder 'BBOX_PAGE1'/'BBOX_OTHER_PAGES' und das PDF-Layout.")
        return None

    # Process and filter extracted article data
    processed_article_data = []
    in_data_section = False # Flag to indicate if we are in the main article data section

    for row in table_data_raw:
        row_joined_upper = " ".join(row).upper()

        # Skip empty rows
        if not any(cell.strip() for cell in row):
            continue

        # Detect the start of the data section (e.g., a row with all underscores)
        if all(cell.strip().startswith('_') for cell in row if cell.strip()):
            in_data_section = True
            continue

        # Skip summary rows or irrelevant headers
        if "PLAN MHD" in row_joined_upper or "SUMME" in row_joined_upper or "GESAMT" in row_joined_upper:
            continue

        # Only process rows if we are past the initial headers and in the data section
        if not in_data_section:
            continue 

        # Extract article number and order quantity from specific columns
        lief_artnr_str = row[0] if len(row) > 0 else ""
        bestellmenge_str = row[1] if len(row) > 1 else "" 

        current_lief_artnr = ""
        # Validate that the article number is numeric
        if lief_artnr_str.strip().isdigit():
            current_lief_artnr = lief_artnr_str.strip()
        else:
            # Skip rows where article number is not valid
            continue

        try:
            # Convert order quantity to float, handling comma as decimal separator
            bestellmenge = float(bestellmenge_str.replace(',', '.').strip())

            # Add valid article data to the list
            if bestelldatum and bestellnummer and lieferdatum and current_lief_artnr and bestellmenge > 0: 
                 processed_article_data.append({
                    "artikelnummer": current_lief_artnr,
                    "bestellmenge": bestellmenge
                })
        except (ValueError, IndexError):
            # Skip rows with invalid quantity format or missing columns
            continue

    if not processed_article_data:
        print(f"WARNUNG: Nach Filterung keine gültigen Artikeldaten für {os.path.basename(pdf_input_path)} gefunden.")
        return None

    # --- CSV File Generation ---
    output_filename = os.path.splitext(os.path.basename(pdf_input_path))[0] + ".csv"
    temp_csv_path = os.path.join(temp_csv_output_dir, output_filename)
    final_csv_path = os.path.join(final_csv_download_dir, output_filename)

    # Special naming convention for HAMMERER files
    if is_hammerer:
        hammerer_date_str = bestelldatum.replace('.', '-') if bestelldatum else ""
        final_csv_path = os.path.join(final_csv_download_dir, f"Bestellformular {hammerer_date_str}.csv")

    print(f"Erstelle temporäre CSV-Datei: {temp_csv_path}")
    with open(temp_csv_path, 'w', newline='', encoding='ISO-8859-1') as csvfile:
        csv_writer = csv.writer(csvfile, delimiter=';')

        # Write header row (e.g., with market password in column 15 for EDEKA)
        header_row_csv = [''] * 15
        header_row_csv[14] = markt_kennwort if not is_hammerer else ""
        csv_writer.writerow(header_row_csv)

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

        # Write article data rows
        for item in processed_article_data:
            row_csv = [''] * 15 
            row_csv[1] = item["artikelnummer"]
            row_csv[4] = str(item["bestellmenge"]).replace('.', ',') # Format quantity with comma for CSV
            csv_writer.writerow(row_csv)

    # Copy the temporary CSV to the final download directory
    try:
        # Copy the temporary CSV to the final destination
        shutil.copy(temp_csv_path, final_csv_path)
        print(f"CSV-Datei erfolgreich erstellt und kopiert nach: {final_csv_path}")

        # Delete the temporary CSV file after successful copy
        os.remove(temp_csv_path)
        print(f"Temporäre Datei '{os.path.basename(temp_csv_path)}' gelöscht.")

        return final_csv_path # Return the path to the successfully created CSV

    except Exception as e:
        print(f"FEHLER beim Kopieren oder Löschen der CSV-Datei: {e}")
        return None

# No main() function here anymore, as this file is imported as a module.