# -*- coding: utf-8 -*-
"""
Created on Fri Feb  7 09:52:52 2025
Modified to handle month overflow (e.g., Dec -> Jan)

@author: gianmaria_c
"""

import tkinter as tk
from tkinter import filedialog
import pdfplumber
from datetime import datetime, timedelta
import uuid

def parse_time(a, t):
    common = [x for x in a if x in t]
    start_idx = -1
    
    l = []
    for idx, item in enumerate(a):
        if item is None: continue # Gestione sicurezza se item è None
        # Strip leading zeros and check if it's '1'
        stripped_num = item.lstrip('0')
        if stripped_num == '1':
            start_idx = idx
            l.append(idx)
            
    # MODIFICA: Prendi tutto dall'inizio del primo '1' fino alla fine della riga valida
    dates = []
    if len(l) >= 1:
        # Invece di tagliare al prossimo '1', prendiamo tutto fino alla fine della lista
        # o finché ci sono numeri validi.
        # Assumiamo che la riga del PDF finisca con i giorni o celle vuote
        raw_dates = a[l[0]:]
        # Puliamo eventuali celle vuote o None alla fine
        dates = [d for d in raw_dates if d and d.strip() != '']
        
    else: 
        print("Something is wrong, please check manually")
        
    return common + dates, l

def parse_shifts(a, s):
    shifts = []
    for x in a: 
        if x in s: 
            shifts.append(x)
        # Aggiunto controllo per None oltre a stringa vuota
        elif x == "" or x is None:
            shifts.append("Empty")
        else:
            print(f"Invalid Shift Found: {x}")
            
    return shifts

def extract_row_from_pdf(pdf_path, search_string):
    allowed_shifts = ["m", "p", "n", "mx", "nx", "Nu", "Mu", "Pu", "MuPu", "Du", "fe", "ffe", "G4.", "G4", "G5.", "G5", "N4", "N5", "co", "m+", "p+", "n+", "af"]
    months = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", 
              "Novembre", "Dicembre"]
    
    extracted_row = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_table()
            if tables:
                for row in tables:
                    # Pulizia base della riga per evitare errori
                    row = ['' if item is None else item for item in row]
                    
                    if len(row) > 1 and row[1] in months:
                        print(f"Found month {row[1]}")
                        time_row = row
                        times, l = parse_time(time_row, months)
                        
                        # Cerchiamo la riga dell'utente NELLA STESSA tabella/pagina
                        # Nota: Questo assume che la riga dei giorni sia sopra la riga dei turni
                        # Se shift_row è in un'altra iterazione, serve logica diversa.
                        # Qui assumiamo che stiamo cercando nella stessa tabella.
                        
                # Riscansioniamo la tabella per trovare la persona, ora che abbiamo i tempi (times)
                
                for sub_row in tables:
                    sub_row = ['' if item is None else item for item in sub_row]
                    if (sub_row[0] == search_string) ^  (sub_row[1] == search_string):  # Match first column (Nome Cognome)
                        # Usiamo l[0] calcolato dalla riga del mese per allinearci
                        if 'l' in locals() and l:
                            extracted_row = sub_row[l[0]:]
                            # Assicuriamoci che extracted_row abbia la stessa lunghezza di dates
                            # (times[1:] contiene i giorni numerici)
                            dates_count = len(times) - 1 # -1 per togliere il nome del mese
                            extracted_row = extracted_row[:dates_count] 
                            
                            shifts = parse_shifts(extracted_row, allowed_shifts)
                            d = dict(times=times, shifts=shifts)
                            return d
                    
    return None  # Return None if no matching row is found

def select_pdf_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    return file_path

def generate_ics_file(shift_dict, output_filename=None, year=2025):
    """
    Generate an ICS file from a dictionary of times and shifts.
    Handles month rollover (e.g. 30, 31, 1, 2).
    """
    # Shift time mappings
    shift_times = {
        'Empty': None,
        'Mu': {'start_time': '08:00', 'duration': timedelta(hours=6)},
        'Du': {'start_time': '08:00', 'duration': timedelta(hours=9)},
        'Pu': {'start_time': '14:00', 'duration': timedelta(hours=6)},
        'Nu': {'start_time': '20:00', 'duration': timedelta(hours=12)},
        'nx': {'start_time': '18:00', 'duration': timedelta(hours=6)},
        'MuPu': {'start_time': '08:00', 'duration': timedelta(hours=12)},
        'm': {'start_time': '08:00', 'duration': timedelta(hours=6)},
        'mx': {'start_time': '12:00', 'duration': timedelta(hours=6)},
        'p': {'start_time': '14:00', 'duration': timedelta(hours=6)},
        'n': {'start_time': '20:00', 'duration': timedelta(hours=12)},
        'G4.': {'start_time': '08:00', 'duration': timedelta(hours=12)},
        'G4': {'start_time': '08:00', 'duration': timedelta(hours=12)},
        'G5.': {'start_time': '08:00', 'duration': timedelta(hours=12)},
        'G5': {'start_time': '08:00', 'duration': timedelta(hours=12)},
        'N4': {'start_time': '20:00', 'duration': timedelta(hours=12)},
        'N5': {'start_time': '20:00', 'duration': timedelta(hours=12)},
        'co': {'start_time': '08:00', 'duration': timedelta(hours=12)}, 
        'm+': {'start_time': '08:00', 'duration': timedelta(hours=6)},
        'p+': {'start_time': '14:00', 'duration': timedelta(hours=6)},
        'n+': {'start_time': '20:00', 'duration': timedelta(hours=12)},
        'af': {'start_time': '08:00', 'duration': timedelta(hours=12)}
    }
    
    # Month mapping
    month_map = {
        'Gennaio': 1, 'Febbraio': 2, 'Marzo': 3, 'Aprile': 4, 'Maggio': 5, 'Giugno': 6, 
        'Luglio': 7, 'Agosto': 8, 'Settembre': 9, 'Ottobre': 10, 'Novembre': 11, 'Dicembre': 12
    }
    
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    
    ics_content = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Custom Shift Calendar//EN'
    ]
    
    # MODIFICA: Gestione intelligente di anno e mese
    start_month_name = shift_dict['times'][0]
    current_month = month_map.get(start_month_name, 1)
    current_year = year
    
    previous_day_val = 0
    
    # Iteriamo su giorni e turni
    # shift_dict['times'][1:] contiene i giorni come stringhe ('1', '2', '31', '1', ecc.)
    days_list = shift_dict['times'][1:]
    shifts_list = shift_dict['shifts']
    shifts_list_no_Empty = [s for s in shifts_list if s != 'Empty']
    print(f"Found {len(days_list)} days and {len(shifts_list_no_Empty)} shifts")
    
    for day_str, shift in zip(days_list, shifts_list):
        if not day_str or day_str == '': continue
        
        try:
            day_val = int(day_str)
        except ValueError:
            continue # Salta se non è un numero
            
        # LOGICA ROLLOVER:
        # Se il giorno corrente è minore del precedente (es. 31 -> 1),
        # siamo passati al mese successivo.
        if day_val < previous_day_val:
            current_month += 1
            if current_month > 12:
                current_month = 1
                current_year += 1
                print(f"Year rollover detected: now {current_year}")
            print(f"Month rollover detected: now month {current_month}")
            
        previous_day_val = day_val
        
        # Skip Empty shifts
        if shift == 'Empty':
            continue
            
        shift_detail = shift_times.get(shift)
        if not shift_detail:
            continue
        
        try:
            # Create datetime for the event
            start_datetime = datetime(current_year, current_month, day_val, 
                                      int(shift_detail['start_time'].split(':')[0]),
                                      int(shift_detail['start_time'].split(':')[1]))
            
            # Calculate end datetime
            end_datetime = start_datetime + shift_detail['duration']
            
            # Create unique identifier
            event_uuid = str(uuid.uuid4())
            
            # Add event to ICS
            ics_content.extend([
                'BEGIN:VEVENT',
                f'UID:{event_uuid}',
                f'DTSTAMP:{dtstamp}',
                f'DTSTART:{start_datetime.strftime("%Y%m%dT%H%M%S")}',
                f'DTEND:{end_datetime.strftime("%Y%m%dT%H%M%S")}',
                f'SUMMARY:{shift}',
                'END:VEVENT'
            ])
        except ValueError as e:
            print(f"Error creating date for {day_val}/{current_month}/{current_year}: {e}")
            continue
    
    ics_content.append('END:VCALENDAR')
    
    ics_string = '\r\n'.join(ics_content)
    
    if output_filename:
        with open(output_filename, 'w') as f:
            f.write(ics_string)
    
    return ics_string

# --- MAIN EXECUTION ---

# Select File
pdf_path = select_pdf_file()

if pdf_path:
    search_string = "Farolfi"  # Inserisci il cognome corretto
    
    try:
        ts = extract_row_from_pdf(pdf_path, search_string)
        
        if ts:
            print(f"Dati estratti per {search_string}")
            # Generate ICS file - Assicurati di passare l'anno corretto se non è il 2025
            ics_content = generate_ics_file(ts, output_filename='shifts_calendar.ics', year=2025)
            print("ICS file generated successfully!")
        else:
            print(f"Nessuna riga trovata per: {search_string}")
            
    except Exception as e:
        print(f"Si è verificato un errore: {e}")
