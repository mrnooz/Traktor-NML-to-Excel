# Script per convertire una collection NML di Traktor in Excel
# Salva questo script in un file .py e assicurati di avere installato pandas e openpyxl

import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl

# Percorso del file NML (es. collection.nml o collection.xml)
input_file = "C:/COLLECTIONbig.nml"  # <--- percorso aggiornato

# Parsing del file XML
print("Caricamento e parsing del file...")
tree = ET.parse(input_file)
root = tree.getroot()

# Cerca tutti gli elementi <ENTRY>
entries = root.findall(".//ENTRY")

# Estrai dati chiave per ogni traccia
data = []
for entry in entries:
    title = entry.attrib.get("TITLE", "")
    artist = entry.attrib.get("ARTIST", "")
    album = entry.attrib.get("ALBUM", "")
    genre = entry.attrib.get("GENRE", "")
    key = entry.attrib.get("KEY", "")
    bpm = ""
    rating = ""
    import_date = ""

    # Cerca sub-elementi
    location = entry.find("LOCATION")
    info = entry.find("INFO")
    tempo = entry.find("TEMPO")
    mod_date = entry.find("MODIFICATION_INFO")

    file_path = location.attrib.get("FILE", "") if location is not None else ""
    bpm = tempo.attrib.get("BPM", "") if tempo is not None else ""
    rating = info.attrib.get("RATING", "") if info is not None else ""
    import_date = mod_date.attrib.get("DATE", "") if mod_date is not None else ""

    # Rimozione caratteri non validi per Excel
    clean_row = {
        "Title": title.replace('=', '').replace('"', '').strip(),
        "Artist": artist.replace('=', '').replace('"', '').strip(),
        "Album": album.replace('=', '').replace('"', '').strip(),
        "Genre": genre.replace('=', '').replace('"', '').strip(),
        "Key": key.replace('=', '').replace('"', '').strip(),
        "BPM": bpm,
        "Rating": rating,
        "Import Date": import_date,
        "File Name": file_path.replace('=', '').replace('"', '').strip()
    }

    data.append(clean_row)

# Crea DataFrame e salva in Excel evitando formule Excel
print("Esportazione in Excel...")
df = pd.DataFrame(data)
output_file = "Traktor_Collection.xlsx"
df.to_excel(output_file, index=False, engine="openpyxl")

print(f"Fatto! File salvato come {output_file}")
