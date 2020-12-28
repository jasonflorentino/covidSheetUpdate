#! Python3

from selenium import webdriver
from time import sleep
import openpyxl
import config
import ezsheets
import os

DRIVER_LOCATION = "/usr/local/bin/chromedriver"
DATA_URL = "https://drive.google.com/file/d/11KF1DuN5tntugNc10ogQDzFnW05ruzLH/view"
XPATH = "/html/body/div[3]/div[3]/div/div[3]/div[2]/div[2]/div[3]"
SPREADSHEET_FILE_NAME = "CityofToronto_COVID-19_Daily_Public_Reporting.xlsx"
SPREADSHEET_FILE_NAME_ALT = "CityofToronto_COVID-19_Daily_Public_Reporting (2).xlsx"
ACTIVE_ROW = 297
ACTIVE_ROW_LINE_NO = 15
PREVIOUS_DATE = '2020-12-27'
PREVIOUS_DATE_LINE_NO = 17

def wait(seconds):
    print(f"Waiting {seconds} seconds...")
    sleep(seconds)

# Download Excel spreadsheet owned by City of Toronto 
while True:
    print("Opening browser...")
    browser = webdriver.Chrome(DRIVER_LOCATION)
    browser.get(DATA_URL)
    wait(4)

    button = browser.find_element_by_xpath(XPATH)
    print(f"Found <{button.tag_name}> element.")

    button.click()
    print(f"Clicked <{button.tag_name}> element.")
    wait(7)

    print("Quitting browser...")
    browser.quit()
    wait(4)

    # Get data from downloaded Excel spreadsheet
    print("Opening downloaded file...")
    try:
        wb = openpyxl.load_workbook(config.DOWNLOAD_FOLDER + SPREADSHEET_FILE_NAME)
        workingPath = config.DOWNLOAD_FOLDER + SPREADSHEET_FILE_NAME
        print(f"Opened {SPREADSHEET_FILE_NAME}.")
    except:
        print(f"Could not find: {SPREADSHEET_FILE_NAME}. Trying alt name...")
        wb = openpyxl.load_workbook(config.DOWNLOAD_FOLDER + SPREADSHEET_FILE_NAME_ALT)
        workingPath = config.DOWNLOAD_FOLDER + SPREADSHEET_FILE_NAME_ALT
        print(f"Opened {SPREADSHEET_FILE_NAME_ALT}.")

    print("Retrieving date...")
    dateSheet = wb['Cases by Reported Date']
    DATE = dateSheet['A2'].value
    DATE = str(DATE).split(" ")[0]

    if DATE == PREVIOUS_DATE:
        print(f"Stopping execution: Spreadsheet data has already been read. (Already read data from: {DATE})")
        print("Deleting downloaded file...")
        os.remove(workingPath)
        break

    print("Retrieving data...")
    sheet = wb['Daily Status']
    TOTAL_CASE_COUNT = sheet['B2'].value
    RECOVERED = sheet['B5'].value
    FATAL = sheet['B6'].value
    CURRENTLY_HOSP = sheet['B8'].value
    CURRENTLY_ICU = sheet['B9'].value

    LATEST_DATA = [DATE, TOTAL_CASE_COUNT, RECOVERED, FATAL, CURRENTLY_HOSP, CURRENTLY_ICU]
    print(f"Success! New data retrieved: {LATEST_DATA}")

    print("Deleting downloaded file...")
    os.remove(workingPath)

    # Update Google Sheet
    print("Accessing Google sheet...")
    doc = ezsheets.Spreadsheet(config.GOOGLE_SHEET_ID)
    sheet = doc[0]
    print(f"Updating next row... (Row: {ACTIVE_ROW})")
    sheet.updateRow(ACTIVE_ROW, LATEST_DATA)

    # Update counters in this script for next time
    print("Reading update.py...")
    content = []
    with open(__file__,"r") as f:
        for line in f:
            content.append(line)

    print("Writing update.py...")
    with open(__file__,"w") as f:
        nextRow = ACTIVE_ROW + 1
        content[ACTIVE_ROW_LINE_NO - 1] = f"ACTIVE_ROW = {nextRow}\n"
        content[PREVIOUS_DATE_LINE_NO - 1] = f"PREVIOUS_DATE = '{DATE}'\n"
        for i in range(len(content)):
            f.write(content[i])

    print("Success! All done!")
    break