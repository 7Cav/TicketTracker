# -*- coding: utf-8 -*-
"""
Created on Sat Apr 27 21:30:51 2019

@author: david
"""

import requests #This is the lib that handled everything web based for python
from bs4 import BeautifulSoup #Literally just used to make things look nice. The find functions are decent too

import gspread #google sheets API use
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *

import math

def SheetGet():
    
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    
    creds = ServiceAccountCredentials.from_json_keyfile_name("Auth.json", scope)
    
    client = gspread.authorize(creds)
    
    sheet = client.open("TicketTracker").sheet1
    
    return sheet

Sheet = SheetGet()

CurrentAOs = ["AdminAccess",
              "Arma3",
              "DCS",
              "Discord",
              "Forums",
              "Other",
              "PostScriptum",
              "Squad"
]

Clerks = ["Jarvis.A",
          "Ratcliff.M",
          "Sweetwater.I",
          "Czar.J",
          "Blackburn.J",
          "Raynor.D",
          "Manus.E",
          "Ticknor.D",
          "Magic",
          "Argus.J",
]


username = input("What is your 7th Cav Email? ")
Password = input("What is your 7th Cav password? ")

    
loopflag = 0 #initial variables for later use

pagenumber = 1
output=[]
PageNumbers=[]

#Used later to trick website into thinking we are using a browser
BrowserHeaders = {'useragent' :"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0"
        }

#Login data needed to, you guessed it, login
login_data = { 
        'login':username,
        'register':	'0',
        'password':	Password,
        'cookie_check':	'1',
        '_xfToken':'',	
        'redirect':	'https://7cav.us/'
}

#Open/crete a csv file
#CSV = comma delimited. Each time there is a comma, there is a new colum in the excel sheet 
f = open("Ticket Tracker.csv", "w")
#start excel formula creation (caution, it's terrible)
f.write(",,,,TICKET COUNTS") #Youll notice all my writes are shifted over, that is because I am lazy and dont want to code a way to ignore ticket data that will be wrote below
f.write("\n")

InitialTicketCells = ["AO","CLERK", "TURN TIME", "", "AOs", 'TICKETS', "AVG TURN TIME", "","CLERK","TOTAL TICKETS","AVG TURN TIME"]

cell_listInitial = Sheet.range("A1:K1")

x = 0
for cells in cell_listInitial:
    cells.value = InitialTicketCells[x]
    x += 1
    
Sheet.update_cells(cell_listInitial, value_input_option='USER_ENTERED')

cell_listTicketTrackerFormuler = Sheet.range("E2:G{}".format(len(CurrentAOs)))

FormulaList = []

FormulaList.append("TOTAL TICKETS")
FormulaList.append("=COUNTA(A:A)-1")
FormulaList.append("=AVERAGE(C:C)")
#Below we will be literally writing in excel formula to take care of data manipulation
f.write(",,,,TOTAL TICKETS,=counta(A:A)-1\n")#Counts total if there is data, -1 to get rid of header
for AO in range(len(CurrentAOs)): #Loop for each AO with formula to count how many tickets are part of it
    f.write(',,,,{},"=COUNTIF(A:A,E{})"\n'.format(CurrentAOs[AO],AO+3)) #.format is great
    
    FormulaList.append('{}'.format(CurrentAOs[AO]))
    FormulaList.append("=COUNTIF(A:A,E{})".format(AO+3))
    FormulaList.append("=AVERAGEIF(A:A,E{},C:C)".format(AO+3))
       
x = 0  
for cells in cell_listTicketTrackerFormuler:
    cells.value = FormulaList[x]

    x += 1
Sheet.update_cells(cell_listTicketTrackerFormuler, value_input_option='USER_ENTERED')

f.write("\n,,,,CLERKS,TOTAL TICKETS,AVG TURN TIME\n")#clerk counting

ClerkList = []

cell_listClerkTrackerFormuler = Sheet.range("I2:K{}".format(len(Clerks)))

for Clerk in range(len(Clerks)):
    f.write(',,,,{},"=COUNTIF(B:B,E{})","=AVERAGEIF(B:B,E{},C:C)"\n'.format(Clerks[Clerk],Clerk+AO+6,Clerk+AO+6))
    
    ClerkList.append('{}'.format(Clerks[Clerk]))
    ClerkList.append('=COUNTIF(B:B,I{})'.format(Clerk+2))
    ClerkList.append('=AVERAGEIF(B:B,I{},C:C)'.format(Clerk+2))

f.write("\n,,,,TOTAL AVG TURN TIME\n") #Overall turn time
f.write(',,,,"=AVERAGE(C:C)"\n')



x = 0  
for cells in cell_listClerkTrackerFormuler:
    cells.value = ClerkList[x]

    x += 1
Sheet.update_cells(cell_listClerkTrackerFormuler, value_input_option='USER_ENTERED')

with requests.Session() as s:
    LoginUrl = 'https://7cav.us/login/login' #Actual login link
    TicketsUrl = 'https://7cav.us/forums/resolved.425/'#Resolved Ticket URL
    
    #Login block below
    r = s.get(LoginUrl, headers=BrowserHeaders) #gets to the webpage with defined header to make it think we are in a browser not a script
    soup = BeautifulSoup(r.content, 'html.parser') #Beautiful soup converts raw text data into html looking format
    login_data['_xfToken'] = soup.find('input', attrs={'name':'_xfToken'})['value']#find login token. For cav website its _xftoken
    r = s.post(LoginUrl, data=login_data, headers=BrowserHeaders)#Post the login command
    #Total resolved ticket pages block below
    Pages = s.get(TicketsUrl, headers=BrowserHeaders) 
    PagesSoup = BeautifulSoup(Pages.content, 'html.parser') 
    PageNumbers = PagesSoup.find('div', attrs={'class':'PageNav'})['data-last']#Last page number is stored in data-last

    #Ticket data grabbing below
    while loopflag != int(PageNumbers):
        ticketurl = TicketsUrl + 'page-' + str(pagenumber)#Creates the url for each page of tickets url takes form of:https://7cav.us/forums/resolved.425/page-[x]
        tickets = s.get(ticketurl,headers=BrowserHeaders)
        Ticketsoup = BeautifulSoup(tickets.content, 'html.parser')
        #Grabs all the data in the input class which holds titles of all threads on page
        #Loops through for however many threads are on page
        for titles in Ticketsoup.find_all('input', attrs={'name': 'threads[]'}):
            RawTickets = (titles.get('title')[29:])#Grabs individual titles and saves it to a variable
            RawTickets = RawTickets.replace('|',',') #Adds commas where | is in thread titles, useful for csv files
            RawTickets = RawTickets.replace('Resolved', '') #Filter out needless characters
            RawTickets = RawTickets.replace('Turn time', '')
            RawTickets = RawTickets.replace('hrs', '')
            RawTickets = RawTickets.replace('hr', '')
            ComaTickets = RawTickets.replace(' ', '')
            output.append(ComaTickets) #Finally append title to end out all title list
            
        loopflag += 1
        pagenumber += 1
x = 0

Result = []

f.write("AO,CLERK,TURN TIME\n")

for x in range(len(output)):
    f.write(output[x]) #write to our CSV, each line is a title 
    f.write("\n")
    Result.append([x.strip() for x in output[x].split(',')]) #Turn singular list in a list with 3 elements on each line
    
f.close()


cell_listA = Sheet.range('A2:C{}'.format(len(Result))) #Cell_list becomes the size of total tickets time 3. This is the amount of cells we will be taking up

fmtheader = cellFormat(
    backgroundColor=color(1, 0.9, 0.9),
    textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1)),
    horizontalAlignment='CENTER')

center = cellFormat(horizontalAlignment='CENTER')

#fmtovertime = cellFormat(backgroundColor = color(1,0,0))
fmtOK = cellFormat(backgroundColor = color(0.13,0.87,0.16))

formatlisting = [('A1:R1', fmtheader)]

Color = [(0.1,1,0.13), (0.25,1,0.09), (0.45,0.99,0.08), (0.65,0.99,0.07), 
         (0.85,0.99,0.05), (0.99,0.91,0.04), (0.98,0.7,0.03), (0.98,0.48,0.02), 
         (0.98,0.26,0.01), (1,0.03,0)]


x = 0
y = 0
for cellA in cell_listA:
    cellA.value = Result[x][y] #Fill in values of each individual cell to use back in the cell list
    if Result[x][2] != 'NF':
        Check = int(Result[x][2])
    if Check >= 24:
        if Check <= 120:
            select = round(Check/12)-1
        else:
            select = 9
        fmt = cellFormat(backgroundColor = color(*Color[select]))
        formatlisting.append(('C{}'.format(x+2), fmt))
        Check = 0
    else:
        formatlisting.append(('C{}'.format(x+2),fmtOK))
    y += 1
    if y == 3: #This is to avoid index out of bounds
        y = 0
        x += 1
        
Sheet.update_cells(cell_listA, value_input_option='USER_ENTERED') #write to google sheets
format_cell_range(Sheet,'A2:R{}'.format(len(Result)), center)
format_cell_ranges(Sheet, formatlisting)






