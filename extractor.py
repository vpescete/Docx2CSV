import docx
import pandas as pd

# Apri il file Word
doc = docx.Document('esempio.docx')

# Parole chiave e parole associate
keywords = {'Nome:': 'Nome', 'Cognome:': 'Cognome', 'Data di nascita:': 'Data di nascita'}

# Dizionario per tenere traccia delle parole associate
parole_associate = {colonna: '' for colonna in keywords.values()}

# Cerca le parole associate nel documento
for paragraph in doc.paragraphs:
    for keyword, associated_word in keywords.items():
        if keyword in paragraph.text:
            indice = paragraph.text.index(keyword) + len(keyword)
            parola_associata = paragraph.text[indice:].strip()
            parole_associate[associated_word] = parola_associata

# Carica il file CSV esistente (se presente)
try:
    df = pd.read_csv('dati.csv', delimiter=';')
except FileNotFoundError:
    df = pd.DataFrame(columns=parole_associate.keys())

# Aggiungi le parole associate come nuova riga
df = pd.concat([df, pd.DataFrame([parole_associate])], ignore_index=True)

# Scrivi i risultati nel file CSV
df.to_csv('dati.csv', index=False, sep=';')
