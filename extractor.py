import os
import shutil
import docx
import pandas as pd

# Imposta le cartelle di origine, destinazione e dati
cartella_origine = 'Files2Scrape'
cartella_destinazione = 'FilesAlreadyScraped'
cartella_dati = 'Data'

# Verifica se la cartella dei dati esiste, altrimenti creala
if not os.path.exists(cartella_dati):
    os.makedirs(cartella_dati)

# Verifica se la cartella di destinazione esiste, altrimenti creala
if not os.path.exists(cartella_destinazione):
    os.makedirs(cartella_destinazione)

# Parole chiave e parole associate
keywords = {'Nome:': 'Nome', 'Cognome:': 'Cognome', 'Data di nascita:': 'Data di nascita', 'Luogo di nascita:': 'Luogo di nascita', 'E-mail:': 'E-mail'}

# Dizionario per tenere traccia delle parole associate
parole_associate = {colonna: '' for colonna in keywords.values()}

# Elabora tutti i file .docx nella cartella di origine
for filename in os.listdir(cartella_origine):
    if filename.endswith('.docx'):
        # Percorso completo del file .docx
        percorso_file = os.path.join(cartella_origine, filename)
        
        # Apri il file Word
        doc = docx.Document(percorso_file)
        
        # Cerca le parole associate nel documento
        for paragraph in doc.paragraphs:
            for keyword, associated_word in keywords.items():
                if keyword in paragraph.text:
                    indice = paragraph.text.index(keyword) + len(keyword)
                    parola_associata = paragraph.text[indice:].strip()
                    parole_associate[associated_word] = parola_associata
        
        # Carica il file CSV esistente (se presente)
        percorso_csv = os.path.join(cartella_dati, 'dati.csv')
        if os.path.exists(percorso_csv):
            df = pd.read_csv(percorso_csv, delimiter=';')
        else:
            df = pd.DataFrame(columns=parole_associate.keys())
        
        # Aggiungi le parole associate come nuova riga
        df = pd.concat([df, pd.DataFrame([parole_associate])], ignore_index=True)
        
        # Scrivi i risultati nel file CSV
        df.to_csv(percorso_csv, index=False, sep=';')
        
        # Sposta il file .docx nella cartella di destinazione
        shutil.move(percorso_file, os.path.join(cartella_destinazione, filename))
