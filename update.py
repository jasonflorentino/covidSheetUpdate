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
FILE_NAME = "CityofToronto_COVID-19_Daily_Public_Reporting"
FILE_EXTENSION = ".xlsx"
ACTIVE_ROW = 298
ACTIVE_ROW_LINE_NO = 15
PREVIOUS_DATE = '2020-12-28'
PREVIOUS_DATE_LINE_NO = 17
INDENT = "  "

def wait(seconds):
    print(f"{INDENT}Waiting {seconds} seconds...")
    sleep(seconds)

# Download Excel spreadsheet owned by City of Toronto using Selenium
while True:
    print("Getting data file...")
    print(f"{INDENT}Opening browser...")
    browser = webdriver.Chrome(DRIVER_LOCATION)
    browser.get(DATA_URL)
    wait(4)

    button = browser.find_element_by_xpath(XPATH)
    print(f"{INDENT}Found <{button.tag_name}> element.")

    button.click()
    print(f"{INDENT}Clicked <{button.tag_name}> element.")
    wait(7)

    print(f"{INDENT}Quitting browser...")
    browser.quit()
    wait(4)

    # Get data from downloaded Excel spreadsheet
    print("Retrieving data from file...")
    print(f"{INDENT}Opening downloaded file...")
    wb = None
    workingPath = ''
    fileAppendNum = 0

    # Try a few potential file names...
    while True:
        if fileAppendNum == 0:
            try:
                fileName = FILE_NAME + FILE_EXTENSION
                print(f"{INDENT}Looking for {fileName}...")
                wb = openpyxl.load_workbook(config.DOWNLOAD_FOLDER + fileName)
                workingPath = config.DOWNLOAD_FOLDER + fileName
                print(f"{INDENT}Opened {fileName}.")
                break
            except:
                print(f"{INDENT}Could not find: {fileName}.\n{INDENT}Trying next name option...")
                fileAppendNum += 1
        else:
            try:
                downloadAppend = f" ({fileAppendNum})"
                fileName = FILE_NAME + downloadAppend + FILE_EXTENSION
                print(f"{INDENT}Looking for {fileName}...")
                wb = openpyxl.load_workbook(config.DOWNLOAD_FOLDER + fileName)
                workingPath = config.DOWNLOAD_FOLDER + fileName
                print(f"{INDENT}Opened {fileName}.")
                break
            except:
                print(f"{INDENT}Could not find: {fileName}.\n{INDENT}Trying next name option...")
                fileAppendNum += 1
        if fileAppendNum == 4:
            print(f"{INDENT}ERROR: Could not find file after {fileAppendNum} tries.")
            break
                
    print(f"{INDENT}Retrieving date...")
    dateSheet = wb['Cases by Reported Date']
    DATE = dateSheet['A2'].value
    DATE = str(DATE).split(" ")[0]

    # Halt if we've seen this data before
    if DATE == PREVIOUS_DATE:
        print(f"{INDENT}Stopping execution: Spreadsheet data has already been read. (Already read data from: {DATE})")
        print(f"{INDENT}Deleting downloaded file...")
        os.remove(workingPath)
        break

    print(f"{INDENT}Retrieving COVID data...")
    sheet = wb['Daily Status']
    TOTAL_CASE_COUNT = sheet['B2'].value
    RECOVERED = sheet['B5'].value
    FATAL = sheet['B6'].value
    CURRENTLY_HOSP = sheet['B8'].value
    CURRENTLY_ICU = sheet['B9'].value

    LATEST_DATA = [DATE, TOTAL_CASE_COUNT, RECOVERED, FATAL, CURRENTLY_HOSP, CURRENTLY_ICU]
    print(f"{INDENT}Success! New data retrieved: {LATEST_DATA}")

    # Delete file once we have the data
    print(f"{INDENT}Deleting downloaded file...")
    os.remove(workingPath)

    # Update Google Sheet
    print("Updating Google Sheet with new data...")
    print(f"{INDENT}Accessing Google sheet...")
    doc = ezsheets.Spreadsheet(config.GOOGLE_SHEET_ID)
    sheet = doc[0]
    print(f"{INDENT}Updating row {ACTIVE_ROW}...")
    sheet.updateRow(ACTIVE_ROW, LATEST_DATA)

    # Update counters in this script for next time
    print("Updating source script...")
    print(f"{INDENT}Reading update.py...")
    content = []
    with open(__file__,"r") as f:
        for line in f:
            content.append(line)

    print("  Writing update.py...")
    with open(__file__,"w") as f:
        nextRow = ACTIVE_ROW + 1
        print(f"{INDENT}Writing: ACTIVE_ROW = {nextRow}")
        content[ACTIVE_ROW_LINE_NO - 1] = f"ACTIVE_ROW = {nextRow}\n"
        print(f"{INDENT}Writing: PREVIOUS_DATE = '{DATE}'")
        content[PREVIOUS_DATE_LINE_NO - 1] = f"PREVIOUS_DATE = '{DATE}'\n"
        for i in range(len(content)):
            f.write(content[i])

    print("Success! All done!")
    break
