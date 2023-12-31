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
    if not os.path.isdir(cartella_destinazione):
        os.makedirs(cartella_destinazione)
    for filename in os.listdir(cartella_sorgente):
        percorso_file = os.path.join(cartella_sorgente, filename)
        if os.path.isfile(percorso_file) and filename.lower().endswith('.docx') and not filename.startswith('~'):
            destinazione_file = os.path.join(cartella_destinazione, filename)
            if not os.path.exists(destinazione_file):
                shutil.copy2(percorso_file, destinazione_file)
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
        if not os.path.exists(cartella_destinazione):
            os.makedirs(cartella_destinazione)
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
frame = wx.Frame(None, title="From .docx to .csv", size=(380, 200))

# Pannello principale
panel = wx.Panel(frame)
# Creazione di un sizer a griglia per le colonne
sizer = wx.GridBagSizer(hgap=6, vgap=6)

# Pulsante per selezionare una cartella
button_cartella = wx.Button(panel, label="Seleziona Cartella")
button_cartella.Bind(wx.EVT_BUTTON, seleziona_cartella)
sizer.Add(button_cartella, pos=(2, 1), flag=wx.ALIGN_CENTER)

# Pulsante per selezionare un singolo file
button_file = wx.Button(panel, label="Seleziona File")
button_file.Bind(wx.EVT_BUTTON, seleziona_file)
sizer.Add(button_file, pos=(4, 1), flag=wx.ALIGN_CENTER)

# Righe separatori
line1 = wx.StaticLine(panel, style=wx.LI_HORIZONTAL)
sizer.Add(line1, pos=(0, 2), span=(7, 0), flag=wx.EXPAND|wx.ALIGN_CENTER_HORIZONTAL)

# Pulsante per eseguire lo script
button_esegui = wx.Button(panel, label="Esegui Script")
button_esegui.Bind(wx.EVT_BUTTON, esegui_script)
sizer.Add(button_esegui, pos=(3, 3), flag=wx.ALIGN_CENTER)

# Righe separatori
line2 = wx.StaticLine(panel, style=wx.LI_HORIZONTAL)
sizer.Add(line2, pos=(0, 4), span=(7, 1), flag=wx.EXPAND|wx.ALIGN_CENTER_HORIZONTAL)

# Pulsante per scaricare il file CSV
button_scarica = wx.Button(panel, label="Scarica CSV")
button_scarica.Bind(wx.EVT_BUTTON, scarica_csv)
sizer.Add(button_scarica, pos=(2, 5), flag=wx.ALIGN_CENTER)

# Pulsante per uscire dall'applicazione
button_esci = wx.Button(panel, label="Quit")
button_esci.Bind(wx.EVT_BUTTON, esci)
sizer.Add(button_esci, pos=(4, 5), flag=wx.ALIGN_CENTER)

# Settaggio del sizer nel pannello
panel.SetSizerAndFit(sizer)

# Visualizza la finestra
frame.Show()

# Avvio dell'event loop dell'applicazione wxPython
app.MainLoop()