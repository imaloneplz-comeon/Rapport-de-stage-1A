import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import fitz
import re
from datetime import datetime
import arabic_reshaper
from bidi.algorithm import get_display
import sys  # Import sys for sys.exit()

# Function to load configuration from config.json
def load_config():
    config_path = r'C:\Users\MSI\Desktop\PDF_Filling_App\config.json'  # Adjust this path as per your actual config location
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file '{config_path}' not found.")
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config

# Loading configuration
try:
    config = load_config()
except FileNotFoundError as e:
    print(f"Error loading configuration: {e}")
    sys.exit(1)  # Exit the script if configuration cannot be loaded

# Define constants from config
TEMPLATES_FOLDER = config['TEMPLATES_FOLDER']
input_pdf_paths = {key: os.path.join(TEMPLATES_FOLDER, value) for key, value in config['input_pdf_paths'].items()}
FONT_PATH = config['FONT_PATH']
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), 'Downloads')

# Data extraction functions
def extract_first_table(file_path):
    with pdfplumber.open(file_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()
        if tables:
            first_table = tables[0]
            return first_table
        else:
            raise ValueError("No tables found in the PDF.")

def extract_second_table(file_path):
    with pdfplumber.open(file_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()
        if len(tables) > 1:
            second_table = tables[1]
            return second_table
        else:
            raise ValueError("Second table not found in the PDF.")

def extract_specific_values_from_first_table(table):
    df = pd.DataFrame(table[1:], columns=table[0])
    date_echeance = df['Date d\'échéance'].iloc[0]
    montant_max_a_placer = df['Montant maximum à\nplacer (en DH)'].iloc[0]
    return date_echeance, montant_max_a_placer

def extract_specific_values_from_second_table(table):
    df = pd.DataFrame(table[1:], columns=table[0])
    tmp_value = df['TMP'].iloc[0]
    return tmp_value

def extract_date_reglement_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()
    date_reglement_pattern = re.compile(r'Date de règlement\s*:\s*(\d{2}/\d{2}/\d{4})')
    date_reglement_match = date_reglement_pattern.search(text)
    date_reglement = date_reglement_match.group(1) if date_reglement_match else 'Date de règlement non trouvée'
    return date_reglement

def format_date(date_str):
    date_obj = datetime.strptime(date_str, "%d/%m/%Y")
    return date_obj.strftime("%d-%m-%y")

def translate_to_arabic(key, value):
    months_ar = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
    if key in ["Date de règlement", "Date d'échéance"]:
        date = datetime.strptime(value, "%d/%m/%Y")
        day = date.day
        month = months_ar[date.month - 1]
        year = date.year
        return f"{day} {month} {year}"
    elif key == "Durée":
        return f"{value.split()[0]} ايام"
    return value

def fill_pdf_francais(input_pdf_path, output_pdf_path, data):
    positions = {
        "Date de règlement": (310, 380),
        "Date d'échéance": (310, 400),
        "Durée": (310, 420),
        "Montant maximum à placer (en DH)": (310, 450),
        "Première valeur dans la colonne TMP": (310, 470)
    }
    doc = fitz.open(input_pdf_path)
    page = doc[0]
    for key, value in data.items():
        position = positions[key]
        page.insert_text(position, value, fontsize=13, fontfile=None, set_simple=True)
        if key == "Date de règlement":
            page.insert_text((384, 261), value, fontsize=14, fontfile=None, set_simple=True)
            page.insert_text((390, 216), value, fontsize=13, fontfile=None, set_simple=True)
    doc.save(output_pdf_path)
    doc.close()

def fill_pdf_arabe(input_pdf_path, output_pdf_path, data):
    positions = {
        "Date de règlement": [(240, 200), (190, 360)],
        "Date d'échéance": [(190, 380)],
        "Durée": [(190, 410)],
        "Montant maximum à placer (en DH)": [(190, 430)],
        "Première valeur dans la colonne TMP": [(190, 450)]
    }
    doc = fitz.open(input_pdf_path)
    page = doc[0]
    for key, value in data.items():
        if key in positions:
            arabic_value = translate_to_arabic(key, value)
            for idx, pos in enumerate(positions[key]):
                reshaped_text = arabic_reshaper.reshape(arabic_value)
                bidi_text = get_display(reshaped_text)
                fontsize = 16 if key == "Date de règlement" and idx == 0 else 13
                page.insert_text(pos, bidi_text, fontsize=fontsize, fontfile=FONT_PATH, set_simple=False)
    doc.save(output_pdf_path)
    doc.close()

def fill_pdf_anglais(input_pdf_path, output_pdf_path, data):
    positions = {
        "Date de règlement": (310, 360),
        "Date d'échéance": (310, 430),
        "Durée": (310, 410),
        "Montant maximum à placer (en DH)": (310, 450),
        "Première valeur dans la colonne TMP": (310, 470)
    }
    additional_positions = [
        (384, 260),
        (310, 380)
    ]
    data["Durée"] = data["Durée"].replace("jours", "days")
    doc = fitz.open(input_pdf_path)
    page = doc[0]
    for key, value in data.items():
        position = positions[key]
        page.insert_text(position, value, fontsize=13, fontfile=None, set_simple=True)
        if key == "Date de règlement":
            for pos in additional_positions:
                page.insert_text(pos, value, fontsize=13, fontfile=None, set_simple=True)
    doc.save(output_pdf_path)
    doc.close()

def get_unique_filename(folder, base_name):
    counter = 1
    filename = f"{base_name}.pdf"
    while os.path.exists(os.path.join(folder, filename)):
        filename = f"{base_name}_{counter}.pdf"
        counter += 1
    return filename

# GUI functions
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        try:
            first_table = extract_first_table(file_path)
            date_echeance, montant_max_a_placer = extract_specific_values_from_first_table(first_table)
            second_table = extract_second_table(file_path)
            tmp_value = extract_specific_values_from_second_table(second_table)
            date_reglement = extract_date_reglement_from_pdf(file_path)
            date_reglement_dt = datetime.strptime(date_reglement, '%d/%m/%Y')
            date_echeance_dt = datetime.strptime(date_echeance, '%d/%m/%Y')
            duration = (date_echeance_dt - date_reglement_dt).days
            data = {
                "Date de règlement": date_reglement,
                "Date d'échéance": date_echeance,
                "Durée": f"{duration} jours",
                "Montant maximum à placer (en DH)": montant_max_a_placer,
                "Première valeur dans la colonne TMP": tmp_value
            }
            for lang, template_filename in config['input_pdf_paths'].items():
                template_path = os.path.join(TEMPLATES_FOLDER, template_filename)
                output_path = os.path.join(DOWNLOADS_FOLDER, get_unique_filename(DOWNLOADS_FOLDER, f"placement_{format_date(date_reglement)}_{lang}"))
                if lang == "francais":
                    fill_pdf_francais(template_path, output_path, data)
                elif lang == "arabe":
                    fill_pdf_arabe(template_path, output_path, data)
                elif lang == "anglais":
                    fill_pdf_anglais(template_path, output_path, data)

            messagebox.showinfo("Success", "PDF files have been filled and saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            sys.exit(1)

def create_gui():
    root = tk.Tk()
    root.title("PDF Filling App")
    root.geometry("300x200")
    upload_button = tk.Button(root, text="Upload PDF", command=upload_file)
    upload_button.pack(pady=50)
    root.mainloop()

# Start the GUI
if __name__ == "__main__":
    create_gui()
