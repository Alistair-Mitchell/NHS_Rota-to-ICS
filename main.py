import pandas as pd
import tabula
import numpy as np
import re
from datetime import datetime, timedelta
import datetime as dt
import matplotlib.pyplot as plt
import matplotlib
from icalendar import Calendar, Event, vCalAddress, vText
import pytz
from collections import namedtuple
import pypdf
import matplotlib as mpl
mpl.rcParams['figure.dpi'] = 300
import os
# Set the JAVA_HOME environment variable to enable Tabula functionality
os.environ["JAVA_HOME"] = "C:\Program Files\Java\jdk-19"

# Define the file path to the PDF that contains the data to be processed
file_path = "E:\Downloads\PdfAnalysis (3).pdf"

# Extract data tables from the second page of the PDF using Tabula
pdf_dataframes = tabula.read_pdf(file_path, pages="all")

# Read the PDF content using PyPDF and extract the text from the first page
pdf_reader = pypdf.PdfReader(file_path)
pdf_text = pdf_reader.pages[0].extract_text()

# Extract the doctor's name from the PDF text (assumes a specific format)
doctor_name = pdf_text.splitlines()[2].split(", ")[0]

# Define banding data for classifying work hours and conditions
banding_data = {
    'Upper bound': [100, 56, 56, 48, 48, 48, 40],
    'Lower bound': [56, 48, 48, 40, 40, 40, 0],
    'Multiplier': [2.0, 1.8, 1.5, 1.5, 1.4, 1.2, 0.0]
}

# Create a pandas DataFrame from the banding data
banding_table = pd.DataFrame(banding_data, index=['3', '2a', '2b', '1a', '1b', '1c', 'nb'])

# Function to calculate the duration between two times
def calculate_duration(start_time, end_time):
    """
    Calculates the duration between two times.
    """
    time_format = '%H:%M'
    time_delta = datetime.strptime(end_time, time_format) - datetime.strptime(start_time, time_format)
    if time_delta.days < 0:
        time_delta = timedelta(days=0, seconds=time_delta.seconds, microseconds=time_delta.microseconds)
    return time_delta

def convert_timedelta_to_hours(duration):
    """
    Converts a timedelta object into hours as a float value.
    - Extracts the number of days and seconds from the timedelta.
    - Converts days to hours and calculates additional hours and minutes from seconds.
    - Returns the total duration in hours.
    """
    days, seconds = duration.days, duration.seconds
    hours = days * 24 + seconds // 3600
    minutes = (seconds % 3600) // 60
    return hours + minutes / 60

def calculate_overlap(start1, end1, start2, end2):
    """
    Calculates the overlap (in hours) between two time intervals.
    - Uses namedtuples for start and end of intervals.
    - Determines the latest start time and the earliest end time.
    - Calculates the overlap duration in hours.
    """
    Range = namedtuple('Range', ['start', 'end'])
    range1 = Range(start1, end1)
    range2 = Range(start2, end2)
    latest_start = max(range1.start, range2.start)
    earliest_end = min(range1.end, range2.end)
    overlap_duration = max(0, convert_timedelta_to_hours(earliest_end - latest_start))
    return overlap_duration

def extract_rota(pdf_dataframes):
    """
    Extracts rota information from PDF dataframes.
    - Processes each sheet and each week in the data.
    - Computes shift details like type, start time, end time, and duration.
    - Identifies weekends and calculates unsocial and overlap hours.
    - Compiles data into a DataFrame and calculates summary statistics.
    """
    rota_entries = []    
    weekend_count = 0
    unsocial_hours_total = 0
    overlap_hours_total = 0

    for sheet_index, sheet in enumerate(pdf_dataframes):  # Iterate through PDF dataframes
        for week_index, row in sheet.iterrows():  # Process each week
            week_data = pd.DataFrame(row)
            week_start_date = week_data.values[0][0]
            dt_week_start = datetime.strptime(week_start_date, '%d %b %Y')  # Parse the start date of the week
            
            for day_index, (day, shift_details) in enumerate(week_data[1:].iterrows()):  # Process each day's shifts
                if shift_details.isnull().any():  # Skip rows with missing values
                    continue

                # Calculate the shift date and extract details
                shift_date = dt_week_start + timedelta(days=day_index)
                shift_month = shift_date.month
                shift_data = shift_details.iloc[0].split("\r")

                shift_type = shift_data[0]  # Shift type (e.g., "Day", "Night")
                start_time = datetime.strptime(shift_date.strftime('%Y-%m-%d') + ' 07:00:00', '%Y-%m-%d %H:%M:%S')
                end_time = datetime.strptime(shift_date.strftime('%Y-%m-%d') + ' 19:00:00', '%Y-%m-%d %H:%M:%S')

                try:
                    # Parse shift start and end times
                    start, end = shift_data[1].split()
                    shift_duration = calculate_duration(start, end)  # Calculate duration of the shift
                    shift_start = datetime.combine(shift_date.date(), datetime.strptime(start, '%H:%M').time())
                    shift_end = shift_start + shift_duration
                except:
                    # Handle errors or missing data by assigning default values
                    shift_start = shift_date
                    shift_end = shift_date + timedelta(days=1)
                    shift_duration = timedelta(0, 0)

                # Format shift details and add to entries
                formatted_date = shift_date.strftime('%a %d %b')
                rota_entry = [formatted_date, shift_type, shift_start, shift_end, shift_duration, week_index + 1, shift_month]
                rota_entries.append(rota_entry)

                # Calculate overlap and update totals
                overlap_hours = calculate_overlap(start_time, end_time, shift_start, shift_end)
                total_shift_hours = convert_timedelta_to_hours(shift_duration)

                if total_shift_hours == 0:
                    continue

                if shift_date.weekday() > 4:  # Check if the shift is during a weekend
                    weekend_count += 1
                    unsocial_hours_total += total_shift_hours

                unsocial_hours_total += abs(overlap_hours - total_shift_hours)
                overlap_hours_total += overlap_hours

    # Compile data into a DataFrame
    rota_df = pd.DataFrame(rota_entries, columns=['Date', 'Shift Type', 'Start Time', 'End Time', 'Duration', 'Week', 'Month'])
    rota_period = rota_df['Start Time'].iloc[-1].date() - rota_df['Start Time'].iloc[0].date()
    total_weeks = (rota_period.days + 1) // 7  # Calculate the total number of weeks
    shift_counts = rota_df['Shift Type'].value_counts()
    total_hours = convert_timedelta_to_hours(rota_df['Duration'].sum())  # Sum up the total hours worked

    try:
        average_weekly_hours = round(total_hours / total_weeks, 1)  # Calculate average weekly hours
    except ZeroDivisionError:
        average_weekly_hours = 0

    try:
        night_shifts = shift_counts.get("Night", 0)  # Count the number of night shifts
    except KeyError:
        night_shifts = 0

    # Print statistics
    print("Total hours:", total_hours)
    print("Night shifts:", night_shifts)
    print("Average weekly hours:", average_weekly_hours)
    print("Weekend days:", weekend_count, "Weekends:", weekend_count // 2)
    print("Unsocial hours:", unsocial_hours_total, f"({round(unsocial_hours_total / total_hours * 100, 2)}% of total)")

    summary = [total_hours, unsocial_hours_total, rota_period.days + 1, night_shifts, weekend_count]
    return summary, rota_df

def export_rota_to_ics(rota, doctor_name):
    """
    Exports the rota information to an ICS calendar file.
    """
    calendar = Calendar()
    calendar.add('prodid', '-//My calendar product//example.com//')
    calendar.add('version', '2.0')

    for _, row in rota.iterrows():
        event = Event()
        event.add('summary', row['Shift Type'])
        event.add('dtstart', row['Start Time'])
        event.add('dtend', row['End Time'])
        event['location'] = vText('Raigmore Hospital,Old Perth Rd, Inverness IV2 3UJ')
        calendar.add_component(event)

    file_name = f"{doctor_name} Rota, {rota['Date'].iloc[0]} - {rota['Date'].iloc[-1]}.ics"

    with open(file_name, 'wb') as file:
        file.write(calendar.to_ical())

    print(file_name, "exported")
    return rota

def band_checker(summary):
    """
    Placeholder for band checking function logic.
    """
    pass

def band_checker2(summary):
    """
    Estimates the band based on summary data.
    """
    total_hours, unsocial_hours, total_days, night_shifts, weekend_days = summary  
    hours_per_day = total_hours / total_days
    band = ""

    while True:
        user_input = input("\nDo you work an on-call rota? (yes/no): ")
        if user_input.lower() in ["yes", "y"]:
            is_on_call = True
            break
        elif user_input.lower() in ["no", "n"]:
            is_on_call = False
            break
        else:
            print("Invalid input. Please enter yes or no.")
    
    match hours_per_day:
        case h if 56 / 7 <= h < 100 / 7:
            band = "3"
        case h if 48 / 7 <= h < 56 / 7:
            band = "2a" if is_on_call else "2b"
        case h if 40 / 7 <= h < 48 / 7:
            band = "1a" if (unsocial_hours / total_hours) > (1 / 3) else "1b"
        case h if 0 / 7 <= h < 40 / 7:
            band = "nb"
        case _:
            print("Error determining band.")

    print("\nEstimated band:")
    print(banding_table.loc[band].to_markdown(), "\n(Note, band 1c not accounted for)")
    
def hours_per_month(rota):
    """
    Calculates and prints the total hours worked for each month.
    """
    for month in rota["Month"].unique():
        month_data = rota.loc[rota["Month"] == month]
        total_hours = convert_timedelta_to_hours(month_data['Duration'].sum())
        print("Month:", str(month) + ",", "Hours worked:", total_hours)

summary_data, rota = extract_rota(pdf_dataframes)
chosen_rota = export_rota_to_ics(rota, doctor_name)
band_checker2(summary_data)

# Band 3	For those working more than 56 hours per week on average or not achieving the required rest. Non-compliant with the new deal, because of excessive hours or other matters.	100%
# Band 2a	For those working between 48 and 56 hours per week on average, most antisocially	80%
# Band 2b	For those working between 48 and 56 hours per week on average, least antisocially	50%
# Band 1a	For those working between 40 and 48 hours per week on average, most antisocially (>1/3)	50%
# Band 1b	For those working between 40 and 48 hours per week on average, moderately antisocially	40%
# Band 1c	For those working between 40 and 48 hours per week on average, least antisocially	20%
# No band	Doctors working on average 40 hours or fewer a week (unless training less than full time)	0% (Doctors in grade FHO1 will receive a 5% uplift)
