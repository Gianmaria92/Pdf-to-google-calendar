# -*- coding: utf-8 -*-
"""
Created on Fri Feb  7 09:52:52 2025

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
        # Strip leading zeros and check if it's '1'
        stripped_num = item.lstrip('0')
        if stripped_num == '1':
            start_idx = idx
            l.append(idx)
            
            
    # Create dates string from that index onwards
    dates = []
    if start_idx != -1 and len(l) == 1:
        dates = a[start_idx:]
        
    elif len(l) == 2:
        print("Found Following month")
        dates = a[start_idx:(l[1]-1)]
        
    else: 
        print("Something is wrong, pleas check manually")
        
    return common+dates, l

def parse_shifts(a, s):
    
    shifts = []
    for x in a: 
        
        if x in s: 
            shifts.append(x)
        elif x == "":
            shifts.append("Empty")
        else:
            print(f"Invalid Shift Found:{x}")
            
    return shifts

def extract_row_from_pdf(pdf_path, search_string):
    allowed_shifts = ["m", "p", "n", "mx", "nx", "Nu", "Mu", "Pu", "MuPu", "Du", "fe", "ffe", "G4.", "G4", "G5.", "G5", "N4", "co"]
    months = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", 
              "Novembre", "Dicembre"]
    
    extracted_row = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_table()
            if tables:
                for row in tables:
                    
                    if row[1] in months:
                        print(f"Found month {row}")
                        time = row
                        times, l = parse_time(time, months)
                    
                    if row and row[1] == search_string:  # Match first column
                        row = ['' if item is None else item for item in row]
                        extracted_row = row[l[0]:]
                        shifts = parse_shifts(extracted_row, allowed_shifts)
                
        d = dict(times=times, shifts=shifts)
        return  d
                    
    return None  # Return None if no matching row is found

def select_pdf_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    return file_path

def generate_ics_file(shift_dict, output_filename=None, year=2025):
    """
    Generate an ICS file from a dictionary of times and shifts.
    
    Parameters:
    shift_dict (dict): Dictionary with 'times' and 'shifts' keys
    output_filename (str, optional): Name of the output .ics file
    year (int, optional): Year for the events, defaults to 2024
    
    Returns:
    str: Content of the generated ICS file
    """
    # Shift time mappings (example values - adjust as needed)
    shift_times = {
        'Empty': None,  # No event
        'Mu': {
            'start_time': '08:00',
            'duration': timedelta(hours=6)
        },
        'Du': {
            'start_time': '8:00',
            'duration': timedelta(hours=9)
        },
        'Pu': {
            'start_time': '14:00',
            'duration': timedelta(hours=6)
        },
        'Nu': {
            'start_time': '20:00',
            'duration': timedelta(hours=12)
        },
        'nx': {
            'start_time': '18:00',
            'duration': timedelta(hours=6)
        },
        'MuPu': {
            'start_time': '08:00',
            'duration': timedelta(hours=12)
        },
        'm': {
            'start_time': '08:00',
            'duration': timedelta(hours=6)
        },
        'mx': {
            'start_time': '12:00',
            'duration': timedelta(hours=6)
        },
        'p': {
            'start_time': '14:00',
            'duration': timedelta(hours=6)
        },
        'n': {
            'start_time': '20:00',
            'duration': timedelta(hours=12)
        },
        'G4.': {
            'start_time': '08:00',
            'duration': timedelta(hours=12)
        },
        'G4': {
            'start_time': '08:00',
            'duration': timedelta(hours=12)
        },
        'G5.': {
            'start_time': '08:00',
            'duration': timedelta(hours=12)
        },
        'G5': {
            'start_time': '08:00',
            'duration': timedelta(hours=12)
        },
        'N4': {
            'start_time': '20:00',
            'duration': timedelta(hours=12)
        },
        'co': {
            'start_time': '08:00',
            'duration': timedelta(hours=12)
        }
    }
    # Month mapping
    month_map = {
        'Gennaio': 1,
        'Febbraio': 2,
        'Marzo': 3, 
        'Aprile': 4, 
        'Maggio': 5, 
        'Giugno': 6, 
        'Luglio': 7, 
        'Agosto': 8, 
        'Settembre': 9, 
        'Ottobre': 10, 
        'Novembre': 11,
        'Dicembre': 12}
    
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    # Prepare ICS content
    ics_content = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Custom Shift Calendar//EN'
    ]
    
    # Process each day
    month = month_map.get(shift_dict['times'][0], 2)  # Default to February
    
    for day, shift in zip(shift_dict['times'][1:], shift_dict['shifts']):
        # Skip Empty shifts
        if shift == 'Empty':
            continue
        
        # Convert day to integer
        day = int(day)
        
        # Get shift details
        shift_detail = shift_times.get(shift)
        if not shift_detail:
            continue
        
        # Create datetime for the event
        start_datetime = datetime(year, month, day, 
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
    
    # Finalize ICS
    ics_content.append('END:VCALENDAR')
    
    # Combine into a single string
    ics_string = '\r\n'.join(ics_content)
    
    # Write to file if filename provided
    if output_filename:
        with open(output_filename, 'w') as f:
            f.write(ics_string)
    
    return ics_string

# Example usage
pdf_path = select_pdf_file()


if pdf_path:
    search_string = "Farolfi"  # Example name to search
    ts = extract_row_from_pdf(pdf_path, search_string)
    print(ts)
    
# Example usage
shift_data = ts

# Generate ICS file
ics_content = generate_ics_file(shift_data, output_filename='shifts_calendar.ics')
print("ICS file generated successfully!")
            
            
            
            
            
            
            
