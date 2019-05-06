# -*- coding: utf-8 -*-
"""
Created on Sat Apr 27 21:30:51 2019

@author: david
"""

import Config

import requests #This is the lib that handled everything web based for python
from bs4 import BeautifulSoup #Literally just used to make things look nice. The find functions are decent too

import os

import gspread #google sheets API use
from oauth2client.service_account import ServiceAccountCredentials
#from gspread_formatting import *
import gspread_formatting

def SheetGet():
    
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    
    currentdir = os.path.dirname(os.path.realpath('DataGrabber_v5.py'))
    
    creds = ServiceAccountCredentials.from_json_keyfile_name(currentdir + "\Auth.json", scope)
    
    client = gspread.authorize(creds)
    
    sheet = client.open("TicketTracker")
    
    WorkSheet = sheet.worksheet("Sheet2")
    
    return WorkSheet

def DataScrape():
    #Used later to trick website into thinking we are using a browser
    BrowserHeaders = {'useragent' :"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0"
            }
    #Login data needed to, you guessed it, login
    login_data = { 
            'login':Config.Username,
            'register':	'0',
            'password':	Config.Password,
            'cookie_check':	'1',
            '_xfToken':'',	
            'redirect':	'https://7cav.us/'
    }
    
    loopflag = 0
    pagenumber = 1
    output = []
    
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
                output.append(ComaTickets + ',' + 'https://7cav.us/threads/' + titles.get('value')) #Finally append title to end out all title list
        
            loopflag += 1
            pagenumber += 1    

    return output

def InitialFormat(cell_listInitial, InitialTicketCells, Sheet, CurrentAOs, Clerks):
    x = 0
    for cells in cell_listInitial:
        cells.value = InitialTicketCells[x]
        x += 1
        
    Sheet.update_cells(cell_listInitial, value_input_option='USER_ENTERED')
    
    cell_listTicketTrackerFormuler = Sheet.range("F2:H{}".format(len(CurrentAOs) + 2))
    
    FormulaList = []
    
    FormulaList.append("TOTAL TICKETS")
    FormulaList.append("=COUNTA(A:A)-1")
    FormulaList.append("=AVERAGE(C:C)")
    #Below we will be literally writing in excel formula to take care of data manipulation
    for AO in range(len(CurrentAOs)): #Loop for each AO with formula to count how many tickets are part of it
        
        FormulaList.append('{}'.format(CurrentAOs[AO]))
        FormulaList.append("=COUNTIF(A:A,F{})".format(AO+3))
        FormulaList.append("=AVERAGEIF(A:A,F{},C:C)".format(AO+3))
           
    x = 0  
    for cells in cell_listTicketTrackerFormuler:
        cells.value = FormulaList[x]
    
        x += 1
    Sheet.update_cells(cell_listTicketTrackerFormuler, value_input_option='USER_ENTERED')
    
    
    ClerkList = []
    
    cell_listClerkTrackerFormuler = Sheet.range("J2:L{}".format(len(Clerks) + 1))
    
    for Clerk in range(len(Clerks)):
        
        ClerkList.append('{}'.format(Clerks[Clerk]))
        ClerkList.append('=COUNTIF(B:B,J{})'.format(Clerk+2))
        ClerkList.append('=AVERAGEIF(B:B,J{},C:C)'.format(Clerk+2))

    
    x = 0  
    for cells in cell_listClerkTrackerFormuler:
        cells.value = ClerkList[x]
    
        x += 1
    Sheet.update_cells(cell_listClerkTrackerFormuler, value_input_option='USER_ENTERED')

def cellWrite(output, Sheet, Color, fmtOK, formatlisting):
    a = 0
    x = 0
    y = 0
    Result = []
   
    for a in range(len(output)):
        Result.append([a.strip() for a in output[a].split(',')]) #Turn singular list in a list with 3 elements on each line
          
    cell_listA = Sheet.range('A2:D{}'.format(len(Result) + 1)) #Cell_list becomes the size of total tickets time 3. This is the amount of cells we will be taking up
    
    for cellA in cell_listA:
        cellA.value = Result[x][y] #Fill in values of each individual cell to use back in the cell list
        if Result[x][2] != 'NF':
            Check = int(Result[x][2])
        if Check >= 24:
            if Check <= 120:
                select = round(Check/12)-1
            else:
                select = 9
            fmt = gspread_formatting.cellFormat(backgroundColor = gspread_formatting.color(*Color[select]))
            formatlisting.append(('C{}'.format(x+2), fmt))
            Check = 0
        else:
            formatlisting.append(('C{}'.format(x+2),fmtOK))
        y += 1
        if y == 4: #This is to avoid index out of bounds
            y = 0
            x += 1
     
    Sheet.update_cells(cell_listA, value_input_option='USER_ENTERED') #write to google sheets
    return Result  

def CSVWriter(CurrentAOs, Clerks, output):
    #Open/crete a csv file
    #CSV = comma delimited. Each time there is a comma, there is a new colum in the excel sheet 
    f = open("Ticket Tracker.csv", "w")
    #start excel formula creation (caution, it's terrible)
    f.write(",,,,TICKET COUNTS") #Youll notice all my writes are shifted over, that is because I am lazy and dont want to code a way to ignore ticket data that will be wrote below
    f.write("\n")
    f.write(",,,,TOTAL TICKETS,=counta(A:A)-1\n")#Counts total if there is data, -1 to get rid of header
    for AO in range(len(CurrentAOs)): #Loop for each AO with formula to count how many tickets are part of it
        f.write(',,,,{},"=COUNTIF(A:A,E{})"\n'.format(CurrentAOs[AO],AO+3)) #.format is great
        
    f.write("\n,,,,CLERKS,TOTAL TICKETS,AVG TURN TIME\n")#clerk counting
    for Clerk in range(len(Clerks)):
        f.write(',,,,{},"=COUNTIF(B:B,E{})","=AVERAGEIF(B:B,E{},C:C)"\n'.format(Clerks[Clerk],Clerk+AO+6,Clerk+AO+6))
        
    f.write("\n,,,,TOTAL AVG TURN TIME\n") #Overall turn time
    f.write(',,,,"=AVERAGE(C:C)"\n')  
    
    f.write("AO,CLERK,TURN TIME\n")
    
    for x in range(len(output)):
        f.write(output[x]) #write to our CSV, each line is a title 
        f.write("\n")
    f.close()

def main():

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
         
    InitialTicketCells = ["AO","CLERK", "TURN TIME", "THREAD LINK", "", "AOs", 'TICKETS', "AVG TURN TIME", "","CLERK","TOTAL TICKETS","AVG TURN TIME"] #Header of google sheet
    
    Sheet = SheetGet()
    
    cell_listInitial = Sheet.range("A1:L1") #setup sheet range
       
    InitialFormat(cell_listInitial, InitialTicketCells, Sheet, CurrentAOs, Clerks) #function call
    
    output = DataScrape() #grab all data we are interested in
    
    fmtheader = gspread_formatting.cellFormat(
        backgroundColor=gspread_formatting.color(0, 0, 0),
        textFormat=gspread_formatting.textFormat(bold=True, foregroundColor=gspread_formatting.color(1, 0.84, 0)),
        horizontalAlignment='CENTER') #formating
        
    fmtOK = gspread_formatting.cellFormat(backgroundColor = gspread_formatting.color(0.13,0.87,0.16)) #formating
    
    formatlisting = [('A1:R1', fmtheader)]
    
    Color = [(0.1,1,0.13), (0.25,1,0.09), (0.45,0.99,0.08), (0.65,0.99,0.07), 
             (0.85,0.99,0.05), (0.99,0.91,0.04), (0.98,0.7,0.03), (0.98,0.48,0.02), 
             (0.98,0.26,0.01), (1,0.03,0)] #color gradient
    
    Result = cellWrite(output, Sheet, Color, fmtOK, formatlisting) #this updates our sheet
        
    gspread_formatting.format_cell_range(Sheet,'A2:R{}'.format(len(Result)), gspread_formatting.cellFormat(horizontalAlignment='CENTER'))
    gspread_formatting.format_cell_ranges(Sheet, formatlisting) #these two lines format the sheet
    
    if Config.CSVOutput == 1 or Config.CSVOutput == 'yes': #if we want csv writer on
        CSVWriter(CurrentAOs, Clerks, output)

if __name__=="__main__":
    main()