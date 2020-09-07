###################################################################################################################
##           EDIT: Indiana updated its data structure 08/24/20, hence this program is no longer valid.           ##
###################################################################################################################

import csv
import requests
from win10toast import ToastNotifier
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyautogui
import webbrowser
import time
from PIL import Image
import pyperclip

toaster = ToastNotifier()

#read filenames
#CHANGE ME BEFORE EVERY RUN!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
histFilename = 'indiana-covid19-data-june-30' #the date at the end of this should match the date on which you run this program
ageFilename = 'june-30-age-group' #the date at the end of this should match the date on which you run this program
last_date = '2020-06-29' #this should match yesterday's date
fillerNum = '98'
date = 'June 30'
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

#set up age-group csv file by creating it and naming first row
ageFile = open(yourDirectory + ageFilename + '.csv', 'a', newline='\n')
ageWriter = csv.writer(ageFile)  
#write column headings          
ageWriter.writerow(['namelsad', 'age_group', 'covid_count'])
ageFile.close()

#set up county-count csv file by creating it and naming first row
historicalFile = open(yourDirectory + histFilename + '.csv', 'a', newline='\n')
historicalWriter = csv.writer(historicalFile)
#write column headings
historicalWriter.writerow(['namelsad10', 'report_date', 'covid_test', 'covid_deaths', 'covid_count_cumsum', 'covid_deaths_cumsum', 'covid_test_cumsum'])
historicalFile.close()

#get data from ISDH back-end
#this is the source data and looks p weird
#if you want to analyze this data, I recommend viewing it on https://jsonformatter.org/json-pretty-print
re = requests.get('https://www.coronavirus.in.gov/map/covid-19-indiana-daily-report-current.topojson')
#tell Python this data is formatted as a JSON
data = re.json()

for key, value in data.items():
    county_data = data['objects']['cb_2015_indiana_county_20m']['geometries']
    
    for county in county_data:
        #print(county['properties']['VIZ_DATE'][28])
        
        for indiv_county in county['properties']['VIZ_DATE']:
            county_name_original = indiv_county['COUNTY_NAME']
            county_name = county_name_original.title() + ' County'
            
            reported_date = indiv_county['DATE']
            covid_test = indiv_county['COVID_TEST']
            covid_deaths = indiv_county['COVID_DEATHS']
            covid_count_cumsum = indiv_county['COVID_COUNT_CUMSUM']
            covid_deaths_cumsum = indiv_county['COVID_DEATHS_CUMSUM']
            covid_test_cumsum = indiv_county['COVID_TEST_CUMSUM']
            
            historicalFile = open(yourDirectory + histFilename + '.csv', 'a', newline='\n')
            historicalWriter = csv.writer(historicalFile)
            historicalWriter.writerow([county_name, reported_date, covid_test, covid_deaths, covid_count_cumsum, covid_deaths_cumsum, covid_test_cumsum])
            historicalFile.close()
            
        for indiv_county in county['properties']['VIZ_D_AGE']:
            
            county_name_original = indiv_county['COUNTY_NAME']
            county_name = county_name_original.title() + ' County'
            age_grp = indiv_county['AGEGRP']
            covid_count = indiv_county['COVID_COUNT']
            
            ageFile = open(yourDirectory + ageFilename + '.csv', 'a', newline='\n')
            ageWriter = csv.writer(ageFile)            
            ageWriter.writerow([county_name, age_grp, covid_count])
            ageFile.close()
            
#Now, filtering to include only rows from last date, then removing duplicates
big_df = pd.read_csv(yourDirectory + histFilename + '.csv')
latest_day = big_df[big_df['report_date']==last_date]
latest_day = latest_day.drop_duplicates()
latest_day.to_excel(yourDirectory + last_date + '.xlsx', index=False)

age_df = pd.read_csv(yourDirectory + ageFilename + '.csv')
age_df = age_df.drop_duplicates()
age_df.to_excel(yourDirectory + ageFilename + '.xlsx', index=False)

#buncha code to update Google Sheets, which in turn powers the CARTO map
scope = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name(insert-your-gspread-creds-here, scope)
client = gspread.authorize(creds)

#Write to county count Google Sheet
date_sheet = client.open(XXX).sheet1
date_sheet.update([latest_day.columns.values.tolist()] + latest_day.values.tolist())

#Write to age groups Google Sheet
age_sheet = client.open(XXX).sheet1
age_sheet.update([age_df.columns.values.tolist()] + age_df.values.tolist())

#toaster simply sends a Windows notification
toaster.show_toast("And we done!", "Python has finished scraping coronavirus data for Indiana and updating Google Sheets. Now on to updating CARTO datasets and filler images.", threaded=True)

#=========================================================================

#Open county count dataset on CARTO
webbrowser.open(XXX)
time.sleep(10)
#Click on the "Sync Now" button to update dataset as on CARTO
pyautogui.click(440, 277)

#stop...wait a minute?
#<Bruno Mars GIF>
time.sleep(10)

#Open age group dataset on CARTO
webbrowser.open(XXX)
time.sleep(10)
#Click on the "Sync Now" button to update dataset as on CARTO
pyautogui.click(440, 277)

#Give it some buffer time in case map also needs time to reflect changes
time.sleep(7)

#==========================================================================

#CARTO link for the main map
webbrowser.open(XXX)

#Factoring in slow webpages, lagging, loading
time.sleep(10)

#take a screenshot of the above page
myScreenshot = pyautogui.screenshot()
myScreenshot.save(yourDirectory + fillerNum + '.jpg')

#open, crop screenshot
filler = Image.open(yourDirectory + fillerNum + '.jpg')

#crop the map out of the screenshot
#coordinates are in a box tuple of form (top-left x, top-left y, bottom-right x, bottom-right y)
croppedFiller = filler.crop((872,277,1270,886))

#copy the cropped map
copyCrop = croppedFiller.copy()

#create a blank image
img = Image.new('RGB',(920,609),(242,246,249))

#paste copy onto new image
img.paste(copyCrop,(261,0))

#save new image to fillers folder
img.save(yourDirectory + fillerNum + '.jpg')

#=============================================================================

indiana_cases = latest_day['covid_count_cumsum'].sum()
indiana_deaths = latest_day['covid_deaths_cumsum'].sum()

#Copy Editor's Note to clipboard for easier access.
#The code below allows for entering time, date (see line #21) and number of cases/deaths formatted with commas
pyperclip.copy('Editor\'s Note: This map will be updated as more information becomes available. This article and map were last updated at ' + time.strftime("%I:%M") + ' p.m. ' + date + ' to report an increase in the numbers of COVID-19 cases and deaths in the state. There are now at least ' + f"{indiana_cases:,d}" + ' cases of the coronavirus in Indiana and ' + f"{indiana_deaths:,d}" + ' deaths. Monroe County reported its first death from the coronavirus April 12.')
toaster.show_toast("And we're done!", "Time to update CEO. Copy is on clipboard.", threaded=True)
