import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tkinter as tk
from tkinter import filedialog

# Fonction pour créer un fichier docx à partir d'une ligne de CSV
def create_docx_from_row(row, filename):
    doc = Document()
    
    for col, value in row.items():
        paragraph = doc.add_paragraph(f'{col}: {value}')
        
        for run in paragraph.runs:
            run.font.size = Pt(20)  # Définit la taille de la police à 20 points
        
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.save(filename)

# Ouvrir une fenêtre de sélection de fichier pour choisir le fichier CSV
def select_csv_file():
    # Créer une fenêtre tkinter
    root = tk.Tk()
    root.withdraw()  # Masquer la fenêtre principale tkinter
    # Ouvrir la boîte de dialogue pour sélectionner un fichier CSV
    file_path = filedialog.askopenfilename(
        title="Sélectionner un fichier CSV",
        filetypes=(("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*"))
    )
    return file_path

csv_file = select_csv_file()

if csv_file:
    df = pd.read_csv(csv_file)

    for index, row in df.iterrows():
        columnNameDoc = 'actions[0].data'
        serialPlate = row[columnNameDoc]  # Récupérer la valeur de la colonne actions[0].data
        # Nom du fichier docx basé sur la valeur de serialPlate
        filename = f'{serialPlate}.docx'
        create_docx_from_row(row, filename)
        print(f'Fichier {filename} créé.')
else:
    print("Aucun fichier CSV sélectionné.")
