import tkinter as tk
from tkinter import font, messagebox, filedialog, ttk, scrolledtext
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
from Assets.py.download_links import download_links, vn_download_links
from Assets.py.muziek_links import muziek_links, vn_muziek_links
from Assets.py.muziekmogrt_links import muziekmogrt_links, vn_muziekmogrt_links
import pymiere
from pymiere import wrappers, objects
import pymiere.exe_utils as pymiere_exe
from pymiere.wrappers import get_system_sequence_presets, time_from_seconds, add_video_track
import sys
import os
from datetime import datetime
import time
import requests
import shutil
import gzip
from lxml import etree
from PIL import Image

########################################################################################################################
#Het huidige versienummer script + aanmaken van log bestanden + verkrijgen download links
########################################################################################################################

# Huidige versie van het script.
versionnr = "1.5"
versionnrdownload_links = vn_download_links
versionnrmuziek_links = vn_muziek_links
versionnrmuziekmogrt_links = vn_muziek_links

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

# Current working directory
cwd = os.getcwd()

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
        # De huidige versie is up to date, laat een notificatie zien aan de gebruiker.
        # Remove # below to show messagebox, but it keeps looping because of a bug.
        # messagebox.showinfo("Up to date", "U gebruikt de nieuwste versie van het script.")
        print("U gebruikt de nieuwste versie van het script.")

# Voer de functie update_to_latest_version uit wanneer het script start.
update_to_latest_version()



########################################################################################################################
# Programmering voor speciale karakters
########################################################################################################################

# Functie om te checken voor speciale karakters
def has_special_characters(input_str):
    # Using regex to check for special characters
    return bool(re.search(r"[!@#$%^&*(),?\":{}|<>_]", input_str))


########################################################################################################################
# Programmering voor progress bar voor downloaden scripts, vormgeving, huisstijlhandboek en vormgeving.
# Maak een dubbele command van progressbar
########################################################################################################################

def progressbar():
    # Maak scherm aan voor progressbar
    progress_window = tk.Toplevel()
    progress_window.title("Bestanden downloaden...")
    progress_window.geometry("350x40")  # Eerste getal is breedte, tweede getal is hoogte

    # Create an indeterminate progress bar
    progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
    progress_bar.pack(ipadx=350, ipady=40)

    # Start the download process
    progress_bar.start(20)  # Start the indeterminate progress bar animation



########################################################################################################################
#Programmering voor downloaden en uitpakken zip file
########################################################################################################################

def download_zip_file(download_link, folder_path):
    logger.info(f"Starting download_zip_file with download_link: {download_link} and folder_path: {folder_path}")
    
    # Define the path for the downloaded zip file
    zip_file = os.path.join(folder_path, "download.zip")
    logger.debug(f"Zip file will be saved to: {zip_file}")
    
    # Download the zip file
    try:
        logger.info("Attempting to download the zip file.")
        response = requests.get(download_link)
        response.raise_for_status()  # Ensure the download was successful
        logger.info("Zip file downloaded successfully.")
    except Exception as e:
        logger.error(f"Error while downloading the zip file: {e}")
        return False

    # Write the zip file to the specified folder
    try:
        logger.info("Attempting to write the downloaded content to the zip file.")
        with open(zip_file, "wb") as f:
            f.write(response.content)
        logger.debug("Zip file written successfully to disk.")
    except Exception as e:
        logger.error(f"Error while writing the zip file to disk: {e}")
        return False

    # Extract the zip file in the specified folder
    try:
        logger.info("Attempting to extract the zip file.")
        with zipfile.ZipFile(zip_file, "r") as zip_ref:
            zip_ref.extractall(folder_path)
        logger.info(f"Zip file extracted successfully to folder: {folder_path}")
    except Exception as e:
        logger.error(f"Error while extracting the zip file: {e}")
        return False

    # Remove the zip file after extraction
    try:
        logger.info(f"Attempting to remove the zip file: {zip_file}")
        os.remove(zip_file)
        logger.info("Zip file removed successfully.")
    except Exception as e:
        logger.error(f"Error while deleting the zip file: {e}")
        return False
    
    logger.info("download_zip_file function completed successfully.")
    return True



########################################################################################################################
#Programmering voor downloaden van een wav bestand.
########################################################################################################################

def download_wav_file(category, muziek_links, folder_path):
    if category not in muziek_links:
        print(f"Category '{category}' not found in muziekmogrt_links.")
        return False

    # Download the mogrt file for the main category
    wav_file_path = os.path.join(folder_path, f"{category}-BUMPER.wav")
    try:
        response = requests.get(muziek_links[category])
        with open(wav_file_path, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded {category} file to {wav_file_path}")
    except Exception as e:
        print(f'Error downloading {category} file: {e}')
        return False

    # Check if the main category is "EnergiekOisterwijk" and download the additional category
    if category == "EnergiekOisterwijk":
        additional_category = "EnergiekOisterwijkExtra"
        if category in muziek_links:
            additional_category_file_path = os.path.join(folder_path, f"{additional_category}-MuziekExtra.wav")
            try:
                response_additional = requests.get(muziek_links[additional_category])
                with open(additional_category_file_path, 'wb') as file_additional:
                    file_additional.write(response_additional.content)
                print(f"Downloaded {additional_category} file to {additional_category_file_path}")
            except Exception as e:
                print(f'Error downloading {additional_category} file: {e}')
                return False



########################################################################################################################
#Programmering voor downloaden van een mogrt bestand.
########################################################################################################################

def download_mogrt_file(category, muziekmogrt_links, folder_path):
    if category not in muziekmogrt_links:
        print(f"Category '{category}' not found in muziekmogrt_links.")
        return False

    # Download the mogrt file for the main category
    mogrt_file_path = os.path.join(folder_path, f"{category}-MuziekBUMPER.mogrt")
    try:
        response = requests.get(muziekmogrt_links[category])
        with open(mogrt_file_path, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded {category} file to {mogrt_file_path}")
    except Exception as e:
        print(f'Error downloading {category} file: {e}')
        return False

    # Check if the main category is "EnergiekOisterwijk" and download the additional category
    if category == "EnergiekOisterwijk":
        additional_category = "EnergiekOisterwijkExtra"
        if category in muziekmogrt_links:
            additional_muziekmogrt_file_path = os.path.join(folder_path, f"{additional_category}-MuziekExtra.mogrt")
            try:
                response_additional = requests.get(muziekmogrt_links[additional_category])
                with open(additional_muziekmogrt_file_path, 'wb') as file_additional:
                    file_additional.write(response_additional.content)
                print(f"Downloaded {additional_category} file to {additional_muziekmogrt_file_path}")
            except Exception as e:
                print(f'Error downloading {additional_category} file: {e}')
                return False

    return True



########################################################################################################################
#Programmering voor voer een titel in + gedeelte thema wijzigen
########################################################################################################################

def run_themescript():
    logger.debug('Starting run_themescript function.')

    # Open a file dialog for the save directory
    folder_path = filedialog.askdirectory()
    logger.debug(f'Folder dialog opened, received path: {folder_path}')

    # Check if the folder path is valid
    if not folder_path:  # User pressed "Escape" or "Cancel"
        logger.info("Folder selection was cancelled or closed. Exiting function.")
        return

    # Get user inputs
    nameeditor = nameeditor_entry.get()
    title = title_entry.get()
    category = selected_category.get()
    logger.debug(f'User inputs - nameeditor: "{nameeditor}", title: "{title}", category: "{category}"')

    # Check for special characters in the input fields
    if has_special_characters(nameeditor) or has_special_characters(title) or has_special_characters(category):
        logger.warning("Special characters detected in input fields.")
        messagebox.showerror("Fout", "Gebruik geen speciale tekens in de invoervelden.")
        return

    # Confirm inputs after validation
    logger.debug(f'Validated inputs - name: {nameeditor}, title: {title}, category: {category}')

    # Ensure required fields are filled before creating the folder
    if title and category:
        try:
            download_link = download_links[category]
            logger.debug(f'Download link for category "{category}": {download_link}')

            # Create unique folder name
            folder_name = f"{today}_{category}_{title}_{nameeditor}"
            folder_name_proj = f"{today}_{category}_{title}"
            folder_path = os.path.join(folder_path, folder_name)
            os.makedirs(folder_path)
            logger.info(f'Created folder at path: {folder_path}')

            # Download and unzip files to the new directory
            download_zip_file(download_link, folder_path)
            logger.info(f'Files downloaded and unzipped to folder: {folder_path}')

            # Rename .prproj files in the "1. PROJECT" subfolder
            project_folder = os.path.join(folder_path, "1. PROJECT")
            renamed = False
            for root, dirs, files in os.walk(project_folder):
                for file in files:
                    if file.endswith(".prproj"):
                        old_name = file
                        new_name_pproj = f"{folder_name_proj}.prproj"
                        os.rename(os.path.join(root, old_name), os.path.join(root, new_name_pproj))
                        logger.debug(f'Renamed project file from "{old_name}" to "{new_name_pproj}"')
                        renamed = True
            if not renamed:
                logger.warning('No .prproj files found in the "1. PROJECT" folder.')

            # Path to the created project file
            pprojfile_path = os.path.join(project_folder, new_name_pproj)
            logger.info(f'Primary project file path: {pprojfile_path}')

            # Process downgrade if requested
            if downgrade_var.get():
                logger.debug('Downgrade option selected.')
                downgraded_path = os.path.join(project_folder, os.path.splitext(new_name_pproj)[0] + "_downgraded.prproj")

                project_data = open_file(pprojfile_path)
                if project_data:
                    logger.debug('Successfully opened project file for downgrade.')
                    converted_data = convert_data(project_data)
                    if converted_data:
                        write_output_file(
                            converted_data,
                            downgraded_path,
                            lambda: messagebox.showinfo(
                                "Downgrade Voltooid",
                                f"Een downgraded project voor Premiere Pro CC 2019 (v36.0) of nieuwer is aangemaakt op de volgende locatie: {downgraded_path}"
                            )
                        )
                        logger.info(f'Downgraded file saved at: {downgraded_path}')
                    else:
                        logger.error('Conversion of project data failed.')
                        messagebox.showerror("Fout", "Conversie van projectgegevens is mislukt.")
                else:
                    logger.error('Failed to open original project file for downgrade.')
                    messagebox.showerror("Fout", "Origineel projectbestand kon niet worden geopend.")
            else:
                logger.debug('Downgrade option not selected.')

            # Notify user of successful creation
            messagebox.showinfo("Map succesvol aangemaakt!", f"Locatie en mapnaam: {folder_path}")
            logger.info('Folder creation and download process completed successfully.')

        except Exception as e:
            logger.exception("An error occurred during the folder creation or file processing.")
            messagebox.showerror("Fout", f"Er is een onverwachte fout opgetreden: {str(e)}")
    else:
        logger.warning("Missing required fields: title or category not provided.")
        messagebox.showerror("Fout", "Voer de naam van de editor/redacteur in, een titel voor de map en selecteer een thema.\n\nControleer of je deze gegevens hebt ingevuld.")

    # Update output field with the path
    output_field.config(state="normal")  # Enable editing
    output_field.delete(1.0, tk.END)  # Clear existing content
    output_field.insert(tk.END, folder_path)  # Insert new folder path
    output_field.config(state="disabled")  # Disable editing again
    logger.debug('Output field updated with the new folder path.')



########################################################################################################################
#Programmering voor voer een titel in + gedeelte thema wijzigen
########################################################################################################################

def run_projectcreator():
    # Log dat het bezig is
    logger.debug('Running script run_projectcreator')
    # Open een venster file dialog voor het selecteren van de save directory
    folder_path = filedialog.askdirectory()

    # Verkrijg de naam, titel en categorie van de invulvelden/keuzeschermen
    nameeditor = nameeditor_entry.get()
    title = title_entry.get()
    category = selected_category.get()
    
    # Check for special characters in the input fields
    if has_special_characters(nameeditor) or has_special_characters(title) or has_special_characters(category):
        messagebox.showerror("Error", "Gebruik geen speciale tekens in de invoervelden.")
        return

    # Log de waardes van de naam, titel en categorie van de invulvelden/keuzeschermen
    logger.debug(f'Obtained values: name={nameeditor}, title={title}, category={category}')

    # Creeër een bestand als de naam, titel en categorie van de invulvelden/keuzeschermen zijn ingevuld.
    if title and category:
        # Verkrijg de download link van de dictionary gebaseerd op de geselecteerde categorie
        download_link = "https://www.dropbox.com/scl/fo/dle2qh4iscs71nbv7b48j/h?rlkey=gukarkdx3unbyijz2558z5koh&dl=1"
        
        # Log the download link
        logger.debug(f'Obtained download link: {download_link}')

        # maak de map met de huidige datum en de opgegeven titel
        folder_name = f"{today}_{category}_{title}_{nameeditor}" # Dit is bedoeld voor de naam van de gecreeërde map
        folder_name_proj = f"{today}_{category}_{title}" # Dit is bedoeld voor de naam van het premiere project
        folder_path = os.path.join(folder_path, folder_name)
        os.makedirs(folder_path)

        # Download the zip file
        download_zip_file(download_link, folder_path)
        #download_wav_file(category, muziek_links, folder_path)
        #download_mogrt_file(category, muziek_links, folder_path)

        # Log de folder path waar de bestanden zijn uitgepakt
        logger.debug(f'Uitpakken bestanden naar: {folder_path}')

        for root, dirs, files in os.walk(f"{folder_path}/1. PROJECT"):
            for file in files:
                if file.endswith(".prproj"):
                    old_name = file
                    new_name_pproj = f"{folder_name_proj}.prproj"
                    os.rename(os.path.join(root, old_name), os.path.join(root, new_name_pproj))

        # Project is in subdirectory 1. PROJECT. We need to navigate to this folder in folder_path
        project_subdirectory = "1. PROJECT"
        pproj_path = os.path.join(folder_path, project_subdirectory)
        pprojfile_path = os.path.join(pproj_path, new_name_pproj)

        if not os.path.isfile(pprojfile_path):
            raise ValueError("Example prproj path does not exists on disk '{}'".format(new_name_pproj))

        # start premiere
        print("Starting Premiere Pro...")
        if pymiere_exe.is_premiere_running()[0]:
            raise ValueError("There already is a running instance of premiere")
        pymiere_exe.start_premiere()

        # open a project
        print("Opening project '{}'".format(new_name_pproj))
        error = None
        # here we have to try multiple times, as on slower systems there is a slight delay between Premiere initialization
        # and when the PymiereLink panel is started and ready to receive commands. Most install will succeed on the first loop
        for x in range(20):
            try:
                pymiere.objects.app.openDocument(pprojfile_path)
            except Exception as error:
                time.sleep(0.5)
            else:
                break
        else:
            raise error or ValueError("Couldn't open path '{}'".format(new_name_pproj))

        # Load MOGRT file into sequence
        sequences = pymiere.objects.app.project.sequences
        sequence = [s for s in sequences if s.name == "LOVO_Leader"]  # search sequence by name
        print("Opening sequence named '{}'".format(sequence))
        if not sequence:
            raise NameError("Something went wrong with {}").format(sequence)
        sequence = sequence[0]
        pymiere.objects.app.project.openSequence(sequence.sequenceID)
        pymiere.objects.app.project.activeSequence = sequence
        
        mogrt_path = os.path.join(pproj_path, "Motion Graphics Template Media", "LOVO Bumperbalk 2023.mogrt")
        print(mogrt_path)
        mgt_clip = sequence.importMGT(  
            path=mogrt_path,  
            time=time_from_seconds(0),  # start time  
            videoTrackIndex=2, audioTrackIndex=0  # on which track to place it  
        )  
        # get component hosting modifiable template properties  
        mgt_component = mgt_clip.getMGTComponent()  
        # handle two types, see Note 2 above
        if mgt_component is None:
            # Premiere Pro type, directly use components
            components = mgt_clip.components
        else:
            # After Effects type, everything is hosted by the MGT component
            components = [mgt_component]

        # Use razor tool to cut the clip in 2
        seq = pymiere.objects.app.project.activeSequence
        time = time_from_seconds(5)
        # format Time object to timecode string
        timecode = time.getFormatted(seq.getSettings().videoFrameRate, seq.getSettings().videoDisplayFormat)
        pymiere.objects.qe.project.getActiveSequence().getVideoTrackAt(2).razor(timecode)
        # Selecte, delete the second clip in the timeline.
        project = pymiere.objects.app.project
        clips = wrappers.list_video(project.activeSequence)
        #clips[1].setSelected(True, True) # Use this to select a clip. 0 is the first clip
        clips[1].remove(True, True)


#        # laat een bericht zien om aan te geven dat het item succesvol is aangemaakt
#        messagebox.showinfo("Map succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
#    else:
#        # Laat een error bericht zien wanneer de naam, titel en/of categorie van de invulvelden/keuzeschermen niet zijn ingevuld.
#        messagebox.showerror("Error", "Voer de naam van de editor/redacteur in, een titel voor de map en selecteer een thema. \n\nControleer of je deze gegevens hebt ingevuld.")
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering venster aanmaken voor converters, hulpmiddelen en vormgeving
########################################################################################################################

progress_bar = None

def open_window_hulpmiddelen():
    # Create a new window
    convert_window = tk.Toplevel()
    convert_window.title("Bestanden converteren")
    convert_window.geometry("380x260")  # Width x Height in pixels

    # Define the font correctly using tk.font.Font
    font = tk.font.Font(family='Arial', size=12)  # Font settings
    root.option_add('*font', font)

    # Explanation label for 'vormgeving downloaden' with ttk.Label
    vormgeving = ttk.Label(convert_window, text="Bij de hulpmiddelen zitten:\n fonts, mogrts, export presets")
    vormgeving.pack(padx=5, pady=5)

    # Button to download 'vormgeving' items using ttk.Button
    vormgeving_button = ttk.Button(convert_window, text="Download hulpmiddelen",
                                   command=lambda: download_overigeitems("Vormgeving", download_links["Vormgeving"]))
    vormgeving_button.pack(pady=5)

    # Explanation label for video conversion with ttk.Label
    convert_label = ttk.Label(convert_window, text="Het volgende is nog in beta!\nHieronder kun je video's converteren\n naar LOVO format")
    convert_label.pack(padx=5, pady=5)

    # Button for converting to LOVO TV preset with ttk.Button
    convert_button_1 = ttk.Button(convert_window, text="Converteer video naar LOVO TV Preset", command=run_convert_script)
    convert_button_1.pack(pady=5)

    # Button for converting to LOVO TV preset + intro with ttk.Button
    convert_button_2 = ttk.Button(convert_window, text="Converteer video naar thema + LOVO TV Preset", command=run_convert_script)
    convert_button_2.pack(pady=5)

def run_convert_script():

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
        def run():
            subprocess.run(command, shell=True, check=True)
            output_field.delete(1.0, tk.END)
            output_field.insert(tk.END, "{}".format(root.filename))
            winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
        t = Thread(target=run)
        t.start()
    except Exception as e:
        # Log the error message to the log file
        logger.exception(e)
        # Show an error message to the user
        messagebox.showerror("Foutmelding", "Er is een fout opgetreden tijdens het converteren van het bestand. Raadpleeg het logbestand voor meer informatie.")

def run_convert_script_with_intro():

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
        def run():
            subprocess.run(command, shell=True, check=True)
            output_field.delete(1.0, tk.END)
            output_field.insert(tk.END, "{}".format(root.filename))
            winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
        t = Thread(target=run)
        t.start()
    except Exception as e:
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

    # Check if the folder path is valid
    if not folder_path:  # User pressed "Escape" or "Cancel"
        logger.info("Folder selection was cancelled or closed. Exiting function.")
        return

    # Create the folder with the current date and the specified title
    folder_name = f"{today}_{item_name}_{current_time}"
    folder_path = os.path.join(folder_path, folder_name)
    os.makedirs(folder_path)

    # Download the zip file
    download_zip_file(download_link, folder_path)

    # Show a message that the item was successfully created
    messagebox.showinfo(f"{item_name} succesvol aangemaakt!", "Locatie en mapnaam is te vinden op: " + folder_path)
    output_field.delete(1.0, tk.END)
    output_field.insert(tk.END, folder_path)



########################################################################################################################
#Programmering knop converteren project naar oudere versie.
########################################################################################################################

# A part of the code is used inside run_themescript
def open_file(path):
    global filename
    tmp = tempfile.mkdtemp()
    filename = os.path.splitext(os.path.basename(path))[0]
    temp_file = tmp + '/' + filename

    if path:  # If a file is chosen
        with gzip.open(path, 'rb') as f_in:
            with open(temp_file, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)
        return etree.parse(temp_file)
    else:
        return None

def convert_data(projectFile):
    global filename
    versionToConvert = "36"  # CC 2019 (v13.0)
    for project in projectFile.xpath("/PremiereData/Project"):
        if project.get('Version'):
            project.set('Version', versionToConvert)
            return etree.tostring(projectFile, encoding="utf-8", pretty_print=True)

def write_output_file(data, output_file, callback):
    with gzip.open(output_file, 'wb') as f:
        f.write(data)
    callback()

########################################################################################################################
#Programmering knop ? rechtsonderin
########################################################################################################################

def run_vraagteken():
    # Log dat de run_vraagteken functie wordt uitgevoerd
    logger.debug('Running run_vraagteken function')
    
    # Laat het info bericht zien
    messagebox.showinfo("Informatie over LOVO MapMaker", f"Product naam: ItemMap_MakerV{versionnr} \nBestandstype: Applicatie \nVersie programma: {versionnr} \nVersie thema projecten: {vn_download_links}\nVersie WAV muziek: {versionnrmuziek_links}\nVersie MOGRT muziek: {vn_muziekmogrt_links}")



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

# Initialize the main window
root = tk.Tk()
root.title("LOVO Map Maker")
root.geometry("400x585")  # Width x Height
font = tk.font.Font(family='Arial', size=12)  # Font settings
root.option_add('*font', font)

# Path to the Azure theme
cwd = os.getcwd()
theme_path = os.path.join(cwd, "Assets", "Theme", "azure.tcl")

try:
    root.tk.call("source", theme_path)
    root.tk.call("set_theme", "dark")
except tk.TclError as e:
    print(f"Error loading Azure theme: {e}")
    print("Falling back to default Tkinter theme.")

# Set window icon
icoon_tskmana = tk.PhotoImage(file="Assets/images/LOGO_LOVO_Zw.png")
root.wm_iconphoto(True, icoon_tskmana)

# Labels and entries
namelabel = ttk.Label(root, text="Voornaam editor:")
namelabel.pack(padx=5, pady=5)

nameeditor_entry = ttk.Entry(root)
nameeditor_entry.pack(padx=5, pady=5)

titlelabel = ttk.Label(root, text="Titel van item:")
titlelabel.pack()

title_entry = ttk.Entry(root)
title_entry.pack(padx=5, pady=(5, 10))

# Category combobox
selected_category = tk.StringVar(root)
selected_category.set("Selecteer thema")  # Default value

categories = ["Amusement", "Archief", "Cultuur", "Natuur", "Nieuws", "Politiek", "Sport", "AanTafelMetClaudy", "EnergiekOisterwijk", "Special"]
category_menu = ttk.Combobox(root, textvariable=selected_category, values=categories)
category_menu.pack(pady=5)

# Buttons with ttk to apply the Azure theme
#create_project_button = ttk.Button(root, text="Creeër project", command=run_projectcreator)
#create_project_button.pack(pady=5)

create_folder_button = ttk.Button(root, text="Maak map aan", command=run_themescript)
create_folder_button.pack(pady=5)

# Checkbox for creating a downgraded version
downgrade_var = tk.BooleanVar()  # Variable to hold checkbox state
downgrade_checkbox = ttk.Checkbutton(root, text="Creeër downgraded Premiere Pro project", variable=downgrade_var)
downgrade_checkbox.pack(pady=5)

# Create a frame to hold the styled Text widget
output_frame = ttk.Frame(root)
output_frame.pack(padx=10, pady=10, fill='x', expand=True)

# Set up the Text widget with read-only mode and styling
output_field = tk.Text(output_frame, width=30, height=5, bg="#737373", fg="#FFFFFF", 
                       wrap='word', relief="flat", font=font, highlightthickness=0,
                       insertbackground="#007acc")  # Replace #007acc with the button color if different
output_field.insert("1.0", "Hier komt de locatie van uw \ngedownloade bestand te staan \nwanneer u dit heeft aangemaakt.")
output_field.config(state="disabled")  # Make the Text widget read-only
output_field.pack(fill='both', expand=True, padx=5, pady=5)


# Additional buttons with ttk
copy_button = ttk.Button(root, text="Open map locatie", command=open_folder)
copy_button.pack(padx=10, pady=10)

question_button = ttk.Button(root, text="?", width=3, command=run_vraagteken)
question_button.pack(side="bottom", anchor="se", padx=5, pady=5)

# Uncomment if you want to add the Handleiding button
# handleiding_button = ttk.Button(root, text="Handleiding", command=download_handleiding)
# handleiding_button.pack(side="bottom", anchor="se", padx=5, pady=5)

huisstijlhandboek_button = ttk.Button(root, text="Download huisstijlhandboek", command=lambda: download_overigeitems("Huisstijlhandboek", download_links["Huisstijlhandboek"]))
huisstijlhandboek_button.pack(side="bottom", anchor="se", padx=5, pady=5)

convert_button = ttk.Button(root, text="Hulpmiddelen", command=open_window_hulpmiddelen)
convert_button.pack(side="bottom", anchor="se", padx=5, pady=5)

root.mainloop()
