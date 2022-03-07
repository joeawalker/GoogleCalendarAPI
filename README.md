# GoogleCalendarAPI
During my time in retail I would recieve work rotas via excel spreadsheets containing the information of who was working which day. This became a repetitive and time consuming task to simply find out my schedule. This also required alarms to be set for each individual day I worked as the start times where never the same. Giving my ambition to automate simple tasks I set out to make a software that meant I could view this information in a much easier and more readable manner.

<br>

## **Extracting Information**

The spreadsheets would all be saved inside a folder labelled "Rotas" that would contain children folders named with the corrisponding years for that set of rotas. The program loops through the spreadsheets, checks to make sure the files contain information, and then retreives all the information in that file for each day. The user can enter the week number that they would like to start searching from so the program doesn't have to search through every single rota every time.

1. First the program will get the date for the day it is looking at and convert it into (yyyy-mm-dd) format
2. Second it will locate my information via my name which is hardcoded into the program as I had no plans to make this into a publicly released application. I have also implemented checks to make sure holidays "hols", empty spaces, or mistypes are not looked at.
3. The program then takes the name, shift and date and passes it through a data formater I wrote that converts it into readable information for the Google Calendar API.
4. After formatting the information the programs is set to gather the information of all other staff members working that day and their shift data too.
5. Giving all this information it is passed through to the API to update my personal Google Calendar events.

<br>

## **Google Calendar API**

I researched into Google's Google Calendar API and created a developer account that would allow me to update my accounts calendar using Python. This required me to generate tokens for validation using pickle and JSON file types. The python script takes the information previously gathered and creates a new event passing the information as different parameters as well as including adding two alarms a day before and an hour and a half before then executes a live update. The console prints out information to display the event created or an error message to say an update has failed.

<p align="center">
  <img width="400" height="300" src="https://raw.githubusercontent.com/joeawalker/GoogleCalendarAPI/main/shifts.JPG">
</p>

<br>

## **Final Results**

The final results can be seen here showing multiple shifts in my calendar and the times of those shifts. I have a widget on my phone that displays my angenda, which is the next handful of events coming up in my calendar, allowing for a much easier method of seeing my work schedule. 

<p align="center">
  <img width="300" height="400" src="https://raw.githubusercontent.com/joeawalker/GoogleCalendarAPI/main/Google%20Agenda.png">
</p>

If an event is selected you can view further details of the names of everyone working that day, the times they are working, location of store and alarms set etc.

<p align="center">
  <img width="300" height="400" src="https://raw.githubusercontent.com/joeawalker/GoogleCalendarAPI/main/Calendar%20Shift.png">
</p>
