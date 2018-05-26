#  This code will read a Tableau Public profile and write information about the vizzes.
#
#  Written by Ken Flerlage, March, 2018
#
#  This code is in the public domain

import sys
import json
import requests
import datetime
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

# Main processing routine.

# Initialize Variables
pageCount = 50 
index = 0
vizCount = 0
matrix = {}
startDate = datetime.date(year=1970, month=1, day=1)

# Open Google Sheet
scope = ['https://spreadsheets.google.com/feeds']

# Read your Google API key from a local json file.
credentials = ServiceAccountCredentials.from_json_keyfile_name('<File Path Here>/creds.json', scope)
gc = gspread.authorize(credentials) # authenticate with Google  

sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/<Spreadsheet ID Here>')
worksheet = sheet.get_worksheet(0)
all_cells = worksheet.range('A1:C6')

foundValid = 1

# Call the Tableau Public API in chunks and write to the Google Sheet.
while (foundValid == 1):
    parameters = {"count": pageCount, "index": index}
    response = requests.get("https://public.tableau.com/profile/api/<Tableau Public Profile Here>/workbooks", params=parameters)
    output = response.json()

    for i in output:
        # Collect viz information.
        title = i['title']
        viewCount = i['viewCount']
        numberOfFavorites = i['numberOfFavorites']
        firstPublishDate = i['firstPublishDate']
        defaultViewRepoUrl = i['defaultViewRepoUrl']
        showInProfile = i['showInProfile']
        viewCount = i['viewCount']

        # Calculations and cleanup of values.
        firstPublishDateFormatted = startDate + datetime.timedelta(milliseconds=firstPublishDate)
        URL ="https://public.tableau.com/views/" + defaultViewRepoUrl.replace("/sheets","") + "?:embed=y&:display_count=yes&:showVizHome=no" 

        # Store all values in an array.
        matrix[vizCount,0]=title
        matrix[vizCount,1]=viewCount
        matrix[vizCount,2]=numberOfFavorites
        matrix[vizCount,3]=str(firstPublishDateFormatted)
        matrix[vizCount,4]=showInProfile
        matrix[vizCount,5]=URL
        vizCount += 1
    
    if not output:
        # We're out of valid vizzes, so quit
        foundValid = 0
    else:
        foundValid = 1
        index += pageCount

# Loop through the matrix and write values for a batch update to Google Sheets.
rangeString = "A2:F" + str(vizCount+1)
cell_list = worksheet.range(rangeString)

row = 0
column = 0

for cell in cell_list: 
    cell.value = matrix[row,column]
    column += 1
    if (column > 5):
        column=0
        row += 1

# Update in batch 
worksheet.update_cells(cell_list)
