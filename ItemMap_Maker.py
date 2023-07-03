import tkinter as tk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from time import sleep
import threading
from threading import Thread
import tkcalendar
import datetime
import os
import sys
import subprocess
import zipfile
import requests
import logging
import functools
import tempfile
import shutil
import winsound
import re
import pymiere
from Assets.download_links import download_links
import pymiere
from pymiere import wrappers
from pymiere import objects
import pymiere.exe_utils as pymiere_exe
from pymiere.wrappers import get_system_sequence_presets
from pymiere.wrappers import time_from_seconds
from pymiere.wrappers import add_video_track
import sys
import os
from datetime import datetime
import time
import requests
import shutil
from PIL import Image

########################################################################################################################
#Het huidige versienummer script + aanmaken van log bestanden + verkrijgen download links
########################################################################################################################

# Huidige versie van het script.
versionnr = "1.0"

# Verkrijg datum in het format YY-MM-DD zonder de -'s
today = datetime.now().strftime("%y%m%d")
current_time = datetime.now().strftime("%H.%M.%S")


# Maak een log bestand aan voor het bijhouden van errors
log_filename = f"LOVOMapMaker_log_{today}_{current_time}.log"
log_file_path = os.path.join('Assets', 'logs', log_filename)
logging.basicConfig(filename=log_file_path, level=logging.DEBUG)
logger = logging.getLogger(__name__)


# Create a logger for the convert script
convert_logger = logging.getLogger('convert_logger')
convert_logger.setLevel(logging.DEBUG)

# Create a file handler for the convert script logger
convert_log_filename = "logProgressConverter.txt"
convert_log_file_path = os.path.join('Assets', 'logs', convert_log_filename)
convert_file_handler = logging.FileHandler(convert_log_file_path)
convert_file_handler.setLevel(logging.DEBUG)

# Create a formatter for the file handler
convert_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
convert_file_handler.setFormatter(convert_formatter)

# Add the file handler to the logger
convert_logger.addHandler(convert_file_handler)


@functools.lru_cache
def get_download_link(category):
    return download_links[category]



########################################################################################################################
#Programmering voor updaten script
########################################################################################################################

def update_to_latest_version():
    # Vervang <USERNAME> en <REPO> met de GitHub gebruikersnaam en repository naam
    api_url = "https://api.github.com/repos/LOIJK/LOVOMapMaker/releases/latest"
    headers = {
        "Accept": "application/vnd.github+json"
    }

    # Vraag de laatste release/update aan via de GitHub API
    response = requests.get(api_url, headers=headers)
    if response.status_code != 200:
        # Er ging iets fout bij het ophalen van de release, log de error en ga terug
        logger.error(f'Error fetching release from GitHub: {response.status_code} {response.text}')
        return

    # Verkrijg tag name van de laatste release (dmv versie_nummer)
    latest_versiontag = response.json()["tag_name"]

    # Remove 'v' from the tag name if it exists
    if latest_versiontag.startswith("v"):
        latest_version = latest_versiontag.replace("v", "", 1)

    # Vergelijk huidige versie nummer tegen die op Github
    if latest_version > versionnr:
        # Er is een nieuwe versie, laat zien aan gebruiker dmv messagebox
        # Remove # below to show messagebox, but it keeps looping because of a bug.
        # messagebox.showinfo("Update to the latest version", "A new version is available. Please wait a moment and then click OK.")
        # Verkrijg de url van de zipfile, hierin zit de laatste versie
        zip_url = response.json()["zipball_url"]
        # Download zip file
        response = requests.get(zip_url)
        if response.status_code != 200:
            # Er ging iets mis bij het downloaden van de zip file, log de error en return
            logger.error(f'Error downloaden van de zip file op GitHub: {response.status_code} {response.text}')
            return
        # Sla de zip file op in een tijdelijke locatie
        zip_file = tempfile.NamedTemporaryFile(delete=False)
        zip_file.write(response.content)
        zip_file.close()
        # Pak het zipbestand uit in een tijdelijke locatie
        temp_dir = tempfile.TemporaryDirectory()
        with zipfile.ZipFile(zip_file.name, "r") as zip_ref:
            zip_ref.extractall(temp_dir.name)
        # Vind het script file terug in de uitgepakte directory
        for root, dirs, files in os.walk(temp_dir.name):
            for file in files:
                if file.endswith(".py"):
                    # Script bestand gevonden, kopieer naar huidige directory
                    shutil.copy2(os.path.join(root, file), ".")
                    # Breek uit de loop en ga uit de functie
                    break
        # Verwijder het tijdelijke zipbestand en de directory
        os.remove(zip_file.name)
        temp_dir.cleanup()
        # Start het script opnieuw om de nieuwe versie te laden
        python = sys.executable
        os.execl(python, python, * sys.argv)
        # Het scriptbestand was niet gevonden in het uitgepakte directory, log een error en return
        logger.error("Scriptbestand niet gevonden in uitgepakte map")
        temp_dir.cleanup()
    else:
        # De huidige version is up to date, laat een notificatie zien aan de gebruiker.
        # Remove # below to show messagebox, but it keeps looping because of a bug.
        # messagebox.showinfo("Up to date", "U gebruikt de nieuwste versie van het script.")
        print("U gebruikt de nieuwste versie van het script.")

# Voer de functie update_to_latest_version uit wanneer het script start.
update_to_latest_version()



########################################################################################################################
#Programmering voor voer een titel in + gedeelte thema wijzigen
########################################################################################################################

def run_themescript():
    # Log dat het bezig is
    logger.debug('Running script')
    # Open een venster file dialog voor het selecteren van de save directory
    folder_path = filedialog.askdirectory()

    # Verkrijg de naam, titel en categorie van de invulvelden/keuzeschermen
    nameeditor = nameeditor_entry.get() 
    title = title_entry.get()
    category = selected_category.get()
    
    # Log de waardes van de naam, titel en categorie van de invulvelden/keuzeschermen
    logger.debug(f'Obtained values: name={nameeditor}, title={title}, category={category}')

    # Creeër een bestand als de naam, titel en categorie van de invulvelden/keuzeschermen zijn ingevuld.
    if title and category:
        # Verkrijg de download link van de dictionary gebaseerd op de geselecteerde categorie
        download_link = download_links[category]
        
        # Log the download link
        logger.debug(f'Obtained download link: {download_link}')

        # maak de map met de huidige datum en de opgegeven titel
        folder_name = f"{today}_{category}_{title}_{nameeditor}" # Dit is bedoeld voor de naam van de gecreeërde map
        folder_name_proj = f"{today}_{category}_{title}" # Dit is bedoeld voor de naam van het premiere project
        folder_path = os.path.join(folder_path, folder_name)
        os.makedirs(folder_path)

        # download het zip-bestand
        zip_file = os.path.join(folder_path, "download.zip")
        try:
            response = requests.get(download_link)
        except Exception as e:
            logger.error(f'Error downloaden zip file: {e}')
            return

        # schrijf het zip-bestand naar de map
        with open(zip_file, "wb") as f:
            f.write(response.content)

        # pak het zip-bestand uit in de map
        try:
            zip_ref = zipfile.ZipFile(zip_file, "r")
            zip_ref.extractall(folder_path)
            zip_ref.close()
        except Exception as e:
            logger.error(f'Error uitpakken zip file: {e}')
            return

        # verwijder het zip-bestand
        os.remove(zip_file)

        for root, dirs, files in os.walk(f"{folder_path}/1. PROJECT"):
            for file in files:
                if file.endswith(".prproj"):
                    old_name = file
                    new_name = f"{folder_name_proj}.prproj"
                    os.rename(os.path.join(root, old_name), os.path.join(root, new_name))

        # Log de folder path waar de bestanden zijn uitgepakt
        logger.debug(f'Uitpakken bestanden naar: {folder_path}')

        # laat een bericht zien om aan te geven dat het item succesvol is aangemaakt
        messagebox.showinfo("Vormgeving succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
    else:
        # Laat een error bericht zien wanneer de naam, titel en/of categorie van de invulvelden/keuzeschermen niet zijn ingevuld.
        messagebox.showerror("Error", "Voer de naam van de editor/redacteur in, een titel voor de map en selecteer een thema. \n\nControleer of je deze gegevens hebt ingevuld.")
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering venster aanmaken + converters
########################################################################################################################

progress_bar = None

def open_window_externevideoconvert():
    global progress_bar
    # Create a new window
    convert_window = tk.Toplevel()
    convert_window.title("Bestanden converteren")
    convert_window.geometry("380x200")
    convert_font = tk.font.Font(family='Arial', size=12)
    convert_window.option_add('*font', convert_font)

    # Create a button for converting to the LOVO TV preset
    convert_button = tk.Button(convert_window, text="Converteer video naar LOVO TV Preset", command=run_convert_script)
    convert_button.pack(pady=5)
    # Create a button for converting to the LOVO TV preset + intro
    convert_button = tk.Button(convert_window, text="Converteer video naar thema + LOVO TV Preset", command=run_convert_script)
    convert_button.pack(pady=5)

def run_convert_script():
    global progress_bar
    try:
        # Open a file dialog to select the file to convert
        root.filename = filedialog.askopenfilename(initialdir = "/", title = "Selecteer een bestand", filetypes = (("alle bestanden", "*.*"), ("MP4 bestanden", "*.mp4"), ("AVI bestanden", "*.avi"), ("MKV bestanden", "*.mkv")))
        # Check if a file was selected
        if not root.filename:
            return
        # Check of het bestand spaties of speciale karakters heeft
        if re.search(" ", root.filename):
            # Show an error message als het bestand spaties of speciale karakters heeft
            messagebox.showerror("Foutmelding", "Het geselecteerde bestand heeft spaties of speciale tekens. Wijzig het en selecteer het bestand opnieuw.")
            return
        
        output_file = "{}_convertLOVOTVPreset.mp4".format(root.filename)
        if os.path.exists(output_file):
            messagebox.showerror("Foutmelding", "Het geconverteerde bestand bestaat al. Verwijder het bestaande bestand of geef een andere naam op.")
            return

        # Define the ffmpeg command to convert the file
        command = "ffmpeg -i {} -vcodec h264 -s 1920x1080 -r 25 -pix_fmt yuv420p -b:v 6M -profile:v main -level 4.1 -color_primaries bt709 -color_trc bt709 -colorspace bt709 -x265-params \"colorprim=bt709:transfer=bt709:colorspace=bt709\" -acodec aac -ar 48000 -ac 2 -b:a 192k -metadata:s:a:0 language=eng -metadata:s:a:0 title=\"Stereo\" -map_metadata 0 -map_chapters 0 -movflags +faststart {}_convertLOVOTVPreset.mp4".format(root.filename, root.filename)
        progress_bar.start()
        def run():
            subprocess.run(command, shell=True, check=True)
            progress_bar.stop()
            output_field.delete(1.0, tk.END)
            output_field.insert(tk.END, "{}".format(root.filename))
            winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
        t = Thread(target=run)
        t.start()
    except Exception as e:
        progress_bar.stop()
        # Log the error message to the log file
        logger.exception(e)
        # Show an error message to the user
        messagebox.showerror("Foutmelding", "Er is een fout opgetreden tijdens het converteren van het bestand. Raadpleeg het logbestand voor meer informatie.")

def run_convert_script_with_intro():
    global progress_bar
    try:
        # Open a file dialog to select the file to convert
        root.filename = filedialog.askopenfilename(initialdir = "/", title = "Selecteer een bestand", filetypes = (("alle bestanden", "*.*"), ("MP4 bestanden", "*.mp4"), ("AVI bestanden", "*.avi"), ("MKV bestanden", "*.mkv")))
        # Check if a file was selected
        if not root.filename:
            return
        # Check of het bestand spaties of speciale karakters heeft
        if re.search(" ", root.filename):
            # Show an error message als het bestand spaties of speciale karakters heeft
            messagebox.showerror("Foutmelding", "Het geselecteerde bestand heeft spaties of speciale tekens. Wijzig het en selecteer het bestand opnieuw.")
            return
        
        output_file = "{}_convertLOVOTVPreset.mp4".format(root.filename)
        if os.path.exists(output_file):
            messagebox.showerror("Foutmelding", "Het geconverteerde bestand bestaat al. Verwijder het bestaande bestand of geef een andere naam op.")
            return
        intro_video = "Assets/ThemaBumper.mov"
        # Define the ffmpeg command to concatenate the intro video and the selected video 
        command = "ffmpeg -i {} -i {} -filter_complex \"[0:v]trim=start=4.15,format=yuva420p[intro];[intro][1:v]blend=all_mode='overlay':all_opacity=0.8[v]\" -map \"[v]\" -c:v libx264 -b:v 6M -r 25 -pix_fmt yuv420p -profile:v main -level 4.1 -color_primaries bt709 -color_trc bt709 -colorspace bt709 -x265-params \"colorprim=bt709:transfer=bt709:colorspace=bt709\" -acodec aac -ar 48000 -ac 2 -b:a 192k -metadata:s:a:0 language=eng -metadata:s:a:0 title=\"Stereo\" -map_metadata 0 -map_chapters 0 -movflags +faststart {}_convertLOVOTVPreset.mp4".format(intro_video, root.filename, root.filename)
        progress_bar.start()
        def run():
            subprocess.run(command, shell=True, check=True)
            progress_bar.stop()
            output_field.delete(1.0, tk.END)
            output_field.insert(tk.END, "{}".format(root.filename))
            winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
        t = Thread(target=run)
        t.start()
    except Exception as e:
        progress_bar.stop()
        # Log the error message to the log file
        logger.exception(e)
        # Show an error message to the user
        messagebox.showerror("Foutmelding", "Er is een fout opgetreden tijdens het converteren van het bestand. Raadpleeg het logbestand voor meer informatie.")



########################################################################################################################
#Programmering knop download vormgeving, handleiding + huisstijlhandboek en evt thema ophalen
########################################################################################################################

def download_overigeitems(item_name, download_link):
    # Log that the download_item function is being executed
    logger.debug(f'Running download_{item_name.lower()} function')

    # Obtain the download link from the dictionary based on the selected theme
    download_link = download_links[item_name]
    
    # Log the download link
    logger.debug(f'Obtained download link: {download_link}')

    # Open a file dialog to select the storage location
    folder_path = filedialog.askdirectory()

    # Create the folder with the current date and the specified title
    folder_name = f"{today}_{item_name}_{current_time}"
    folder_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_path)

    # Download the zip file
    zip_file = os.path.join(folder_path, "download.zip")
    try:
        response = requests.get(download_link)
    except Exception as e:
        logger.error(f'Error downloading zip file: {e}')
        return

    # Write the zip file to the folder
    with open(zip_file, "wb") as f:
        f.write(response.content)

    # Extract the zip file to the folder
    try:
        zip_ref = zipfile.ZipFile(zip_file, "r")
        zip_ref.extractall(folder_path)
        zip_ref.close()
    except Exception as e:
        logger.error(f'Error extracting zip file: {e}')
        return

    # Delete the zip file
    os.remove(zip_file)

    # Log the folder where the files are extracted to
    logger.debug(f'Extracting files to: {folder_path}')

    # Show a message that the item was successfully created
    messagebox.showinfo(f"{item_name} succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering knop ? rechtsonderin
########################################################################################################################

def run_vraagteken():
    # Log dat de run_vraagteken functie wordt uitgevoerd
    logger.debug('Running run_vraagteken function')
    
    # Laat het info bericht zien
    messagebox.showinfo("Informatie over LOVO MapMaker", f"Product naam: ItemMap_MakerV{versionnr} \nBestandstype: Applicatie \nVersie nummer: {versionnr}")



########################################################################################################################
#Programmering 'Copy' knop
########################################################################################################################

# Create function to copy text
def open_folder():
    # Get the text from the output field
    text = output_field.get(1.0, tk.END)
    
    # Remove the trailing newline character
    folder_location = text[:-1]
    
    # Open the folder in the file explorer
    os.startfile(folder_location)



########################################################################################################################
#Programmering GUI
########################################################################################################################

root = tk.Tk()
root.title("LOVO Map Maker")
root.geometry("400x665") # Eerste getal is breedte, tweede getal is hoogte
font = tk.font.Font(family='Arial', size=12) # Semibold werkt niet, bold niet gebruiken
root.option_add('*font', font)

icoon_tskmana = tk.PhotoImage(file="Assets/images/LOGO_LOVO_Zw.png") # maak een fotoimage-object van het icoonbestand
root.wm_iconphoto(True, icoon_tskmana) # stel het icoon van het hoofdvenster in

namelabel = tk.Label(root, text="Voornaam editor:")
namelabel.pack(padx=5, pady=5)

nameeditor_entry = tk.Entry(root)
nameeditor_entry.pack(padx=5, pady=5)

titlelabel = tk.Label(root, text="Titel van item:")
titlelabel.pack()

title_entry = tk.Entry(root)
title_entry.pack(padx=5, pady=5)

selected_category = tk.StringVar(root)
selected_category.set("Selecteer thema") # instellen als standaardwaarde

categories = ["Amusement", "Archief", "Cultuur", "Natuur", "Nieuws", "Politiek", "Sport"]
category_menu = ttk.Combobox(root, textvariable=selected_category, values=categories)
category_menu.pack(pady=5)

button = tk.Button(root, text="Maak map aan", command=run_themescript)
button.pack(pady=5)

# Maak een knop met de naam "Vormgeving"
label = tk.Label(root, text="Overige hulpmiddelen:")
label.pack(pady=(30, 5))

# Maak een knop om de vormgeving te downloaden
vormgeving_button = tk.Button(root, text="Download vormgeving", command=lambda: download_overigeitems("Vormgeving", download_links["Vormgeving"]))
vormgeving_button.pack(pady=5)
vormgeving_button.pack(pady=5)

# Maak een text veld aan dat o.a. het net aangemaakte folder_path kan laten zien.
output_field = tk.Text(root, width=30, height=5)
output_field.insert(tk.END, "Hier komt de locatie van uw \ngedownloade bestand te staan \nwanneer u dit heeft aangemaakt.")
output_field.pack(padx=10, pady=10)

# Knop om de map locatie over te nemen
copy_button = tk.Button(root, text="Open map locatie", command=open_folder)
copy_button.pack(padx=10, pady=10)

# Converteer venster
convert_button = tk.Button(root, text="Video's converteren", command=open_window_externevideoconvert)
convert_button.pack(pady=5)

# Create a progress bar
progress_bar = ttk.Progressbar(root, mode='determinate',length=275)
progress_bar.pack(pady=20)

button = tk.Button(root, text="?", width=3, height=1, command=run_vraagteken)
button.pack(side="bottom", anchor="se", padx= 5, pady=5)

# Tot nadere informatie is de knop handleiding uit en niet zichtbaar. Bij het verwijderen van de # zal de knop weer zichtbaar zijn.
#handleiding_button = tk.Button(root, text="Handleiding", command=download_handleiding)
#handleiding_button.pack(side="bottom", anchor="se", padx= 5, pady=5)

huisstijlhandboek_button = tk.Button(root, text="Download huisstijlhandboek", command=lambda: download_overigeitems("Huisstijlhandboek", download_links["Huisstijlhandboek"]))
huisstijlhandboek_button.pack(side="bottom", anchor="se", padx= 5, pady=5)

root.mainloop()
