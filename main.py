import pandas as pd
import tabula
import numpy as np
import re
from datetime import datetime,timedelta
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
os.environ["JAVA_HOME"] ="C:\Program Files\Java\jdk-19"

# %matplotlib auto
# %matplotlib inline
file = "E:\Downloads\PdfAnalysis (2).pdf"


dfs = tabula.read_pdf(file, pages="2")

reader = pypdf.PdfReader(file)
text = reader.pages[0].extract_text()
name = text.splitlines()[2].split(", ")[0]

year = 2025

data = {'Upper bound': [100,  56,56,48,48,48,40],
        'Lower bound': [56,48,48,40,40,40,   0],
        'Multiplier': [2.0, 1.8,1.5,1.5,1.4, 1.2, 0.0]}

# Creates pandas DataFrame.
banding = pd.DataFrame(data, index=['3','2a','2b','1a','1b','1c','nb'])


def dura(s1,s2):
    FMT = '%H:%M'
    tdelta = datetime.strptime(s2, FMT) - datetime.strptime(s1, FMT)
    if tdelta.days < 0:
        tdelta = timedelta(
            days=0,
            seconds=tdelta.seconds,
            microseconds=tdelta.microseconds
            )
    return tdelta

def convert_timedelta(duration):
    days, seconds = duration.days, duration.seconds
    # print(days, seconds)
    hours = days * 24 + seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = (seconds % 60)
    hours_2 = minutes/60
    hours = hours+hours_2
    return hours

def date_time_overlap(s1,e1,s2,e2):
    Range = namedtuple('Range', ['start', 'end'])
    r1 = Range(s1, e1)
    # print("R1",r1)
    r2 = Range(s2, e2)
    # print("R2",r2)
    latest_start = max(r1.start, r2.start)
    earliest_end = min(r1.end, r2.end)
    # earliest_end = min(r2.end, r2.end)
    delta = (earliest_end - latest_start)
    delta = convert_timedelta(delta)
    # print("delta",delta)
    overlap = max(0, delta)
    # print(overlap)
    # print(overlap)
    return overlap

def get_rota(dfs):
    basket = []    
    i = 0
    overlap = 0
    antilap = 0
    for pdf_sheet,extracted_frame  in enumerate(dfs) :        
        for week_index, row in extracted_frame.iterrows():
            week_frame = pd.DataFrame(row)
            week_start = week_frame.values[0][0]
            dt_week_start = datetime.strptime(week_start, '%d %b %Y')
            for day_index, [day, shiftcell] in enumerate(week_frame[1:].iterrows()):
                if shiftcell.isnull().any():
                    continue
                shiftdate = dt_week_start+timedelta(day_index)
                shiftmonth = shiftdate.month
                g =shiftcell.iloc[0].split("\r")

                shift = g[0]
                
                start_check_time = dt.time(7,0,0)
                end_check_time = dt.time(19,0,0)
                start_check_date=shiftdate.combine(
                    shiftdate.replace(year=year).date(), 
                    start_check_time
                    )
                end_check_date=shiftdate.combine(
                    shiftdate.replace(year=year).date(), 
                    end_check_time
                    )
                
                try:
                    start,end =g[1].split()
                    duration = dura(start,end)
                    start_dt=shiftdate.combine(
                        shiftdate.date(), 
                        datetime.strptime(start, '%H:%M').time()
                        )
                    end_dt = start_dt+duration
                
                except:
                    start_dt = shiftdate
                    end_dt = shiftdate+ timedelta(days=1)
                    duration = timedelta(0,0)
                    
                date = shiftdate.strftime('%a %d %b')
                jam = [date,shift,start_dt,end_dt,duration,week_index+1,shiftmonth]
                basket.append(jam)
                # print(jam)
                
                uhoh=date_time_overlap(start_check_date,end_check_date,start_dt,end_dt)
                dur_con = convert_timedelta(duration)
                if dur_con ==0:
                    continue
                # print(dur_con,uhoh)
                # print(date,abs(uhoh-dur_con),"antisocial hours")
                if (shiftdate.weekday() >4):
                    # print(date)
                    i+=1
                    antilap = antilap+dur_con
                antilap= antilap + abs(uhoh-dur_con)
                overlap = overlap + uhoh
                
                
        rota = pd.DataFrame(basket,columns=['hr_date', 'Shift','Start','End','Duration','Week','Month'])
        k = rota['Start'].iloc[-1].date()-rota['Start'].iloc[0].date()
        weeks = (k.days+1)//7
        counts = rota['Shift'].value_counts()
        hours = rota["Duration"].sum()
            
        hours=convert_timedelta(hours)
        
        
        
        try:
            average_hours = round(hours/weeks,1)
        except ZeroDivisionError:
            result = 0
        try:
            night = counts["Night"]
        except KeyError:
            night = 0  
            
        
        print("They work",hours, "total hours and does",night,"night shifts.")
        print("They work an average of",average_hours,"hours per week over",weeks,"weeks")
        print("They work",i,"weekend days, or",i//2,"weekends in",weeks,"weeks. This is an average of 1 in",round(weeks/(i//2),2))
        print("They work a total of",antilap,"unsocial hours, which equates to",str(round(antilap/hours,4)*100)+"% of total","\n")
        # band_checker2(hours,k.days+1,antilap)
        summary = [hours,antilap,k.days+1,night,i,weeks]
        return (summary,rota)
    
def rota_to_ics(rota,Doctor_Name):
    wr = rota
    cal = Calendar()
    cal.add('prodid', '-//My calendar product//example.com//')
    cal.add('version', '2.0')
    for index,row in wr.iterrows():
        event = Event()
        event.add('summary', row['Shift'])
        event.add('dtstart', row['Start'])
        event.add('dtend', row['End'])
        event['location'] = vText('Raigmore Hospital,Old Perth Rd, Inverness IV2 3UJ')
        
        cal.add_component(event)

    name = f"{Doctor_Name} Rota, {wr['hr_date'].iloc[0]} - {wr['hr_date'].iloc[-1]}.ics"

    f = open(name, 'wb')
    f.write(cal.to_ical())
    f.close()
    # print("\n")
    print(name,"exported")
    return wr
    
def band_checker(wr):
    for week  in range(wr['Week'].iloc[-1]):
        working_week =  wr.loc[wr["Week"]==week]
        # print(working_week)
        
        # jam = [date,shift,start_dt,end_dt,duration,week_sheet+1]
        # basket.append(jam)
    # summary = pd.DataFrame(basket,columns=['Total hours', 'Unsocial hours','Weekend work?','Unsocial proportion','weeks','Week'])
def band_checker2(summary):
    hours,unsocial,days,nights,weekends,weeks= summary  
    hpd = hours/days
    band = ""
    while True:
        user_input = input("\nDo you work an on call rota: ")
        if user_input.lower() in ["yes", "y"]:
            OC = True
            break
        elif user_input.lower() in ["no", "n"]:
            OC = False
            break
        else:
            print("Invalid input. Please enter yes/no.")
    print(hpd)
    match hpd:
        case a if 56/7 <= a <  100/7:
            band = "3"  
        case b if 48/7 <= b <  56/7:

            match OC:
                case True:
                    band = "2a or 2b"
                case False:
                    band = "2a or 2b"
        case v if 40/7 <= v <  48/7:
            match (unsocial/hours)>(1/3):
                case True:
                    band = "1a"
                case False:
                    band = "1b"
        case d if 0/7 <= d <  40/7:
            band = "nb"
        case _:
            print("uhoh")
    print("\nEstimated band:")
    print(banding.loc[band].to_markdown(),"\n(Note, band 1c not accounted for)")
    

def hours_per_month(wr):
    for x in wr["Month"].unique():
        
        egg = wr.loc[wr["Month"]==x]
        egg2 = convert_timedelta(egg['Duration'].sum(axis=0))
        print("Month:",str(x)+",","Hours worked:",egg2)

summary_data,rota=get_rota(dfs)
chosen_rota = rota_to_ics(rota,name)
band_checker2(summary_data)

# Band 3	For those working more than 56 hours per week on average or not achieving the required rest. Non-compliant with the new deal, because of excessive hours or other matters.	100%
# Band 2a	For those working between 48 and 56 hours per week on average, most antisocially	80%
# Band 2b	For those working between 48 and 56 hours per week on average, least antisocially	50%
# Band 1a	For those working between 40 and 48 hours per week on average, most antisocially (>1/3)	50%
# Band 1b	For those working between 40 and 48 hours per week on average, moderately antisocially	40%
# Band 1c	For those working between 40 and 48 hours per week on average, least antisocially	20%
# No band	Doctors working on average 40 hours or fewer a week (unless training less than full time)	0% (Doctors in grade FHO1 will receive a 5% uplift)