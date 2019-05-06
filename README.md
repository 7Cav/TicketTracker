# 7Cav IMO Ticket Tracker

This ReadMe is for the DataGrabber script used to pull
raw thread title data from the s6 resolved tickets
forum on the 7th Cavalry forum. https://7cav.us/

------------------------------------------------------
                                  
REQUIRES PYTHON 3.7+ TO BE INSTALLED ON YOUR COMPUTER                                        
When running the script it will ask you for your email
and password that you use for the 7th Cavalry website.
This information is not stored in any fashion, and is 
cleared by python's garabge collector when the script
finishes. 
The script will output a .csv file in the same path
which will have the raw ticket data pulled from the 
forum.                                                                                

https://docs.google.com/spreadsheets/d/1A0Cec7Kuvy8ziK68m0jcufaFswDX7olVLkeqIYUwSCI/edit?usp=sharing

Modules needed:
requests,
bs4,
gspread,
oauth2client,
gspread_formatting
