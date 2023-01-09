import tkinter as tk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
import datetime
import os
import sys
import zipfile
import requests
import logging
import functools
import tempfile
import shutil
from Assets.download_links import download_links
sys.path.append(os.path.join(os.getcwd(), "Assets")) # Voor het verkrijgen van het bestand download_links.py

########################################################################################################################
#Het huidige versienummer script + aanmaken van log bestanden + verkrijgen download links
########################################################################################################################

# Huidige versie van het script.
versionnr = "1.0"

# Verkrijg datum in het format YY-MM-DD zonder de -'s
today = datetime.datetime.now().strftime("%y%m%d")
current_time = datetime.datetime.now().strftime("%H.%M.%S")

# Maak een log bestand aan voor het bijhouden van errors
log_filename = f"LOVOMapMaker_log_{today}_{current_time}.log"
log_file_path = os.path.join('Assets', 'logs', log_filename)
logging.basicConfig(filename=log_file_path, level=logging.DEBUG)
logger = logging.getLogger(__name__)

@functools.lru_cache
def get_download_link(category):
    return download_links[category]



########################################################################################################################
#Programmering voor updaten script
########################################################################################################################

def update_to_latest_version():
    # Vervang <USERNAME> en <REPO> met de GitHub gebruikersnaam en repository naaam
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
    latest_version = response.json()["tag_name"]
    # Vergelijk huidige versie nummer tegen die op Github
    if latest_version > versionnr:
        # Er is een nieuwe versie, laat zien aan gebruiker dmv messagebox
        messagebox.showinfo("Updaten naar de laatste versie", f"Er is een nieuwe versie beschikbaar. Wacht heel even en druk dan op ok.")
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
        messagebox.showinfo("Up to date", "U gebruikt de nieuwste versie van het script.")

# Voer de funtie: update_to_latest_version uit wanneer het script start.
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

        # Locatie waar de map staat
        current_dir = os.getcwd()
        folder_path = os.path.join(current_dir, folder_name)

        # Log de folder path waar de bestanden zijn uitgepakt
        logger.debug(f'Uitpakken files to: {folder_path}')

        # laat een bericht zien om aan te geven dat het item succesvol is aangemaakt
        messagebox.showinfo("Vormgeving succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
    else:
        # Laat een error bericht zien wanneer de naam, titel en/of categorie van de invulvelden/keuzeschermen niet zijn ingevuld.
        messagebox.showerror("Error", "Voer de naam van de editor/redacteur in, een titel voor de map en selecteer een thema. \n\nControleer of je deze gegevens hebt ingevuld.")
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering knop download vormgeving + thema ophalen
########################################################################################################################

def download_vormgeving():
    # Log dat de download_vormgeving functie wordt uitgevoerd
    logger.debug('Running download_vormgeving function')

    # Haal de downloadlink op uit de woordenboek op basis van het geselecteerde thema
    download_link = download_links["Vormgeving"]
    
    # Log de downloadlink
    logger.debug(f'Obtained download link: {download_link}')

    # Open een bestandsdialoog om de opslagmap te selecteren
    folder_path = filedialog.askdirectory()

    # Maak de map met de huidige datum en de opgegeven titel aan
    folder_name = f"{today}_Vormgeving_{current_time}"
    folder_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_path)

    # Download het zip-bestand
    zip_file = os.path.join(folder_path, "download.zip")
    try:
        response = requests.get(download_link)
    except Exception as e:
        logger.error(f'Error downloading zip file: {e}')
        return

    # Schrijf het zip-bestand naar de map
    with open(zip_file, "wb") as f:
        f.write(response.content)

    # Pak het zip-bestand uit in de map
    try:
        zip_ref = zipfile.ZipFile(zip_file, "r")
        zip_ref.extractall(folder_path)
        zip_ref.close()
    except Exception as e:
        logger.error(f'Error extracting zip file: {e}')
        return

    # Verwijder het zip-bestand
    os.remove(zip_file)

    # Locatie waar de map staat
    current_dir = os.getcwd()
    folder_path = os.path.join(current_dir, folder_name)

    # Log de map waar de bestanden naar uitgepakt worden
    logger.debug(f'Extracting files to: {folder_path}')

    # Geef een melding dat het item succesvol is aangemaakt
    messagebox.showinfo("Vormgeving succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering knop download handleiding
########################################################################################################################

def download_handleiding():
    # Log dat de download_handleiding functie wordt uitgevoerd
    logger.debug('Running download_handleiding function')

    # Haal de downloadlink op uit de woordenboek op basis van het geselecteerde thema
    download_link = download_links["Handleiding"]
    
    # Log de downloadlink
    logger.debug(f'Obtained download link: {download_link}')

    # Open een file dialof om de opslagmap te selecteren
    folder_path = filedialog.askdirectory()

    # Maak de map met de huidige datum en de opgegeven titel aan
    folder_name = f"{today}_Handleiding_{current_time}"
    folder_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_path)

    # Download het zip-bestand
    zip_file = os.path.join(folder_path, "download.zip")
    try:
        response = requests.get(download_link)
    except Exception as e:
        logger.error(f'Error downloading zip file: {e}')
        return

    # Schrijf het zip-bestand naar de map
    with open(zip_file, "wb") as f:
        f.write(response.content)

    # Pak het zip-bestand uit in de map
    try:
        zip_ref = zipfile.ZipFile(zip_file, "r")
        zip_ref.extractall(folder_path)
        zip_ref.close()
    except Exception as e:
        logger.error(f'Error extracting zip file: {e}')
        return

    # Verwijder het zip-bestand
    os.remove(zip_file)

    # Locatie waar de map staat
    current_dir = os.getcwd()
    folder_path = os.path.join(current_dir, folder_name)

    # Log de map waar de bestanden naar uitgepakt worden
    logger.debug(f'Extracting files to: {folder_path}')

    # Geef een melding dat het item succesvol is aangemaakt
    messagebox.showinfo("Handleiding succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering knop download huisstijlhandboek
########################################################################################################################

def download_huisstijlhandboek():
    # Log dat de download_huisstijlhandboek functie wordt uitgevoerd
    logger.debug('Running download_huisstijlhandboek function')

    # Haal de downloadlink op uit de woordenboek op basis van het geselecteerde thema
    download_link = download_links["Huisstijlhandboek"]
    
    # Log de downloadlink
    logger.debug(f'Obtained download link: {download_link}')

    # Open een bestandsdialoog om de opslagmap te selecteren
    folder_path = filedialog.askdirectory()

    # Maak de map met de huidige datum en de opgegeven titel aan
    folder_name = f"{today}_Huisstijlhandboek_{current_time}"
    folder_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_path)

    # Download het zip-bestand
    zip_file = os.path.join(folder_path, "download.zip")
    try:
        response = requests.get(download_link)
    except Exception as e:
        logger.error(f'Error downloading zip file: {e}')
        return

    # Schrijf het zip-bestand naar de map
    with open(zip_file, "wb") as f:
        f.write(response.content)

    # Pak het zip-bestand uit in de map
    try:
        zip_ref = zipfile.ZipFile(zip_file, "r")
        zip_ref.extractall(folder_path)
        zip_ref.close()
    except Exception as e:
        logger.error(f'Error extracting zip file: {e}')
        return

    # Verwijder het zip-bestand
    os.remove(zip_file)

    # Locatie waar de map staat
    current_dir = os.getcwd()
    folder_path = os.path.join(current_dir, folder_name)

    # Log de map waar de bestanden naar uitgepakt worden
    logger.debug(f'Extracting files to: {folder_path}')

    # Geef een melding dat het item succesvol is aangemaakt
    messagebox.showinfo("Huisstijlhandboek succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
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
def copy_text():
    # Get the text from the output field
    text = output_field.get(1.0, tk.END)
    
    # Remove the trailing newline character
    text = text[:-1]
    
    # Copy the text to the clipboard
    root.clipboard_append(text)



########################################################################################################################
#Programmering GUI
########################################################################################################################

root = tk.Tk()
root.title("LOVO Map Maker")
root.geometry("400x600") # Eerste getal is breedte, tweede getal is hoogte
font = tk.font.Font(family='Arial', size=12) # Semibold werkt niet, bold niet gebruiken
root.option_add('*font', font)

icoon_tskmana = tk.PhotoImage(file="Assets/images/LOGO_LOVO_Zw.gif") # maak een fotoimage-object van het icoonbestand
root.wm_iconphoto(True, icoon_tskmana) # stel het icoon van het hoofdvenster in

namelabel = tk.Label(root, text="Voer de naam van de editor/redacteur in:")
namelabel.pack(padx=5, pady=5)

nameeditor_entry = tk.Entry(root)
nameeditor_entry.pack(padx=5, pady=5)

titlelabel = tk.Label(root, text="Voer een titel voor de map in:")
titlelabel.pack()

title_entry = tk.Entry(root)
title_entry.pack(padx=5, pady=5)

selected_category = tk.StringVar(root)
selected_category.set("Klik hier om het thema te wijzigen") # instellen als standaardwaarde

categories = ["Amusement", "Archief", "Cultuur", "Natuur", "Nieuws", "Politiek", "Sport"]
category_menu = tk.OptionMenu(root, selected_category, *categories)
category_menu.pack(pady=5)

button = tk.Button(root, text="Start project maker", command=run_themescript)
button.pack(pady=5)

# Maak een knop met de naam "Vormgeving"
label = tk.Label(root, text="Overige hulpmiddelen:")
label.pack(pady=(30, 5))

# Maak een knop om de vormgeving te downloaden
vormgeving_button = tk.Button(root, text="Download vormgeving en font", command=download_vormgeving)
vormgeving_button.pack(pady=5)

# Maak een text veld aan dat o.a. het net aangemaakte folder_path kan laten zien.
output_field = tk.Text(root, width=30, height=5)
output_field.insert(tk.END, "Hier komt de locatie van uw \ngedownloade bestand te staan \nwanneer u dit heeft aangemaakt.")
output_field.pack(padx=10, pady=10)

# Create copy button
copy_button = tk.Button(root, text="Kopieer maplocatie", command=copy_text)
copy_button.pack(padx=10, pady=10)

button = tk.Button(root, text="?", width=3, height=1, command=run_vraagteken)
button.pack(side="bottom", anchor="se", padx= 5, pady=5)

handleiding_button = tk.Button(root, text="Handleiding", command=download_handleiding)
handleiding_button.pack(side="bottom", anchor="se", padx= 5, pady=5)

handleiding_button = tk.Button(root, text="Huisstijlhandboek", command=download_huisstijlhandboek)
handleiding_button.pack(side="bottom", anchor="se", padx= 5, pady=5)

root.mainloop()
