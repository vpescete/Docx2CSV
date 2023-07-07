import os
import shutil
import docx
import pandas as pd
import wx
import wx.lib.filebrowsebutton as filebrowse
import subprocess


os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Funzione per copiare i file .docx dalla cartella selezionata nella cartella Files2Scrape
def copia_files(cartella_sorgente):
    cartella_destinazione = 'Files2Scrape'
    for filename in os.listdir(cartella_sorgente):
        percorso_file = os.path.join(cartella_sorgente, filename)
        if os.path.isfile(percorso_file) and filename.lower().endswith('.docx') and not filename.startswith('~'):
            shutil.copy(percorso_file, cartella_destinazione)
            print(f"Il file {filename} è stato copiato nella cartella Files2Scrape.")
    print("Copia dei file .docx completata.")

# Funzione per selezionare una cartella
def seleziona_cartella(event):
    dialog = wx.DirDialog(None, "Seleziona una cartella", style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
    if dialog.ShowModal() == wx.ID_OK:
        cartella_scelta = dialog.GetPath()
        copia_files(cartella_scelta)
        print("I file .docx sono stati copiati nella cartella Files2Scrape.")
    dialog.Destroy()

# Funzione per selezionare uno o più file
def seleziona_file(event):
    dialog = wx.FileDialog(None, "Seleziona uno o più file", wildcard="Documenti Word (*.docx)|*.docx", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE)
    if dialog.ShowModal() == wx.ID_OK:
        percorsi_file = dialog.GetPaths()
        cartella_destinazione = 'Files2Scrape'
        for percorso in percorsi_file:
            shutil.copy(percorso, cartella_destinazione)
        print("I file .docx sono stati copiati nella cartella Files2Scrape.")
    dialog.Destroy()

# Funzione per eseguire lo script Python
def esegui_script(event):
    # Esegue il file extractor.py
    subprocess.call(["python3", "extractor.py"])

# Funzione per scaricare il file dati.csv
def scarica_csv(event):
    dialog = wx.DirDialog(None, "Seleziona la cartella di destinazione", style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
    if dialog.ShowModal() == wx.ID_OK:
        cartella_destinazione = dialog.GetPath()
        nome_file = 'dati_nuovo.csv'
        percorso_file = os.path.join('Data', 'dati.csv')
        shutil.copy(percorso_file, os.path.join(cartella_destinazione, nome_file))
    dialog.Destroy()

# Funzione per uscire dall'applicazione
def esci(event):
    app.ExitMainLoop()

# Creazione dell'applicazione wxPython
app = wx.App()

# Creazione della finestra principale
frame = wx.Frame(None, title="Programma con wxPython", size=(400, 200))

# Pannello principale
panel = wx.Panel(frame)

# Pulsante per selezionare una cartella
button_cartella = wx.Button(panel, label="Seleziona Cartella")
button_cartella.Bind(wx.EVT_BUTTON, seleziona_cartella)

# Pulsante per selezionare uno o più file
button_file = wx.Button(panel, label="Seleziona File")
button_file.Bind(wx.EVT_BUTTON, seleziona_file)

# Pulsante per l'esecuzione dello script
button_esegui = wx.Button(panel, label="Esegui Script")
button_esegui.Bind(wx.EVT_BUTTON, esegui_script)

# Pulsante per il download del file CSV
button_scarica = wx.Button(panel, label="Scarica CSV")
button_scarica.Bind(wx.EVT_BUTTON, scarica_csv)

# Pulsante per uscire dall'applicazione
button_esci = wx.Button(panel, label="Quit")
button_esci.Bind(wx.EVT_BUTTON, esci)

# Layout del pannello
sizer = wx.BoxSizer(wx.VERTICAL)
sizer.Add(button_cartella, 0, wx.ALL, 10)
sizer.Add(button_file, 0, wx.ALL, 10)
sizer.Add(button_esegui, 0, wx.ALL, 10)
sizer.Add(button_scarica, 0, wx.ALL, 10)
sizer.Add(button_esci, 0, wx.ALL, 10)
panel.SetSizer(sizer)

# Visualizza la finestra
frame.Show()

# Avvio dell'event loop dell'applicazione wxPython
app.MainLoop()