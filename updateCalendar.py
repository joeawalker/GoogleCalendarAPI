import os
import getpass
import pandas as pd
import numpy as np
import regex as re
from google_auth_oauthlib.flow import InstalledAppFlow
from apiclient.discovery import build
import pickle

#Rotas are in folders sorted by year, sets the year that is required to be accessed
year = 2021

#Sets location to find rotas and changes directory to access files
loc = ("/Users/" + str(getpass.getuser()) + "/OneDrive - Robert Gordon University/Rotas/"+str(year))
os.chdir(loc)

#Function to update google calendar using the Google Calendar API
def calendar(year, month, day, startH, startM, end, staff):
    scopes = ["https://www.googleapis.com/auth/calendar"]

    #Loads credentials via tokens for authentication
    flow = InstalledAppFlow.from_client_secrets_file("/Users/" + str(getpass.getuser()) + "/Desktop/Work/Code/client_secret.json", scopes=scopes)
    credentials = pickle.load(open("/Users/" + str(getpass.getuser()) + "/Desktop/Work/Code/token.pkl", "rb"))
    pickle.dump(credentials, open("/Users/" + str(getpass.getuser()) + "/Desktop/Work/Code/token.pkl", "wb"))

    service = build("calendar", "v3", credentials=credentials)

    result = service.calendarList().list().execute()

    calendar_id = result['items'][0]['id']

    from datetime import datetime, timedelta

    start_time = datetime(int(year), int(month), int(day), int(startH), int(startM), 0) #Assigns data passed through function to single variable
    end_time = start_time + timedelta(hours=int(end))
    time_zone = 'Europe/London'
    print("\n"+str(start_time) +"\n"+ str(end_time))

    #The event to be added to the calendar
    event = {
        'summary': 'Work',
        'location': 'Fossil',
        'description': staff, #Adds the list of staff and their shifts to the description section of the calendar event
        'start': {
            'dateTime': start_time.strftime("%Y-%m-%dT%H:%M:%S"), #Sets when the event starts
            'timeZone': time_zone,
        },
        'end': {
            'dateTime': end_time.strftime("%Y-%m-%dT%H:%M:%S"), #Sets when the event ends
            'timeZone': time_zone,
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
            {'method': 'popup', 'minutes': 24 * 60}, #Sets reminder of event 24 hours before event
            {'method': 'popup', 'minutes': 90}, #Sets reminder of event an hour and a half before event
            ],
        },
        'event': {
            "background": "red"
        }
    }

    try:
        service.events().insert(calendarId=calendar_id, body=event).execute()     #Executes update
    except:
        print("Could not update calendar")

#Formats data for calander API to read
def format_data(data, dayOfWeek):
    data = data.split(":") #Splits data into accessible sections [name, shift, date]

    date = data[2].split("-")
    year = date[0].split(" ")[1]
    month = date[1]
    day = date[2]

    shift = data[1].strip()
    shift = shift.split("-")
    hour = shift[0][:2]
    minute = shift[0][2:4]

    length = (int(shift[1])-int(shift[0]))/100 #Gets length of shift by subtracting end time from start time and deviding by 100

    staff = ""

    #Gets other staff members that are working that day and their shifts
    for i in range(len(df[dayOfWeek])):
        if len(str(df[dayOfWeek][i]))==9:
            name = df['Name'][i]
            name = name.strip(' ')
            shift = df[dayOfWeek][i]
            person=name+": "+shift
            staff+=person+"\n"

    calendar(year, month, day, hour, minute, length, staff) #Passes information into the calendar API function

    # print("\nYear: "+str(year),   #Prints information sent to calendar
    # "\nMonth: "+str(month),
    # "\nDay: "+str(day),
    # "\nStart Hour: "+str(hour),
    # "\nStart Minute: "+str(minute),
    # "\nShift Length: "+str(length),
    # "\n#  Staff  #\n"+str(staff))

#Function to retrieve shift information
def getShiftInfo(day):
    #Declares list of shifts
    shiftList = []
    #Declares who to search for
    name = "Joe"

    #Retrieves date (yyyy-mm-dd) from dataframe column under the day of the week (eg. Sunday)
    date = df[day]
    date = date.values.tolist() #Converts series into list
    date = str(date[0])
    date = date.split(" ")
    date = date[0] #Final extraction of date (yyyy-mm-dd)

    #Gets shift of detail of person
    person = df.loc[df['Name'] == name] #Locates person by name
    shift = person[day] #Finds shift in relation to day
    shift = shift.fillna('0') #Changes empty cells to have 0's
    shift = shift.values.tolist() #Converts series into list

    #Checks for real shift via "-" and adds to shiftList
    if "-" in shift[0]:
        shiftList.append(shift[0])

    #Prints shifts (name: start time - end time : yyyy-mm-dd)
    for shift in shiftList:
        data = name+": "+shift + " : " + date
        format_data(data, day)

#Function to retrieve rotas in an entire week
def getWeek():
    week = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for day in week:
        getShiftInfo(day)

#Function to retreive dataframe for a weeks rota
def run(week):
    for file in os.listdir(loc): #Loops through files in rota directory
        if file.endswith(".xlsx"): #Checks if file is excel spreadsheet
            if "~" not in file and week in file: #Checks if rota is not open and if the week in question is available
                df = pd.read_excel(file, skiprows=[0, 1, 2, 3, 4]) #Skips useless rows
                df = df.replace(np.nan, '', regex=True)
                return df
        
#Funtion used to find latest week number
def checkNum():
    highNum = 0 #Delcares initial highest number
    for file in os.listdir(loc): #Loops through files in the loc path
        if file.endswith(".xlsx"): #Check if file is an excel document
            testNum = re.findall(r'\d+', str(file)) #Extracts numbers from string and inserts them into a list
            testNum = testNum[0]
            testNum = int(testNum)
            if testNum > highNum: #Checks the newest number is greater than the previous value
                highNum = testNum
    return highNum


maxWeek = checkNum() #Assigns value of the most recent week to maxWeek
week = input("What week would you like to start from: ")

while int(week) <= maxWeek: #Loops through selected week through to most recent week
    df = run(week) #Retrieves dataframe
    if df is None: #If df returns as NoneType then rota not available
        print("\n# Week "+str(week)+" not available #")
        week = int(week) + 1
        week = str(week)   
    else:  
        getWeek() #Retreives rota week data
        week = int(week) + 1
        week = str(week)

ex = input("\nFinished")