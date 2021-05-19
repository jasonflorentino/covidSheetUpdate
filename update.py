#! Python3

from selenium import webdriver
from time import sleep
from os import remove as os_remove
from os import rename as os_rename
import datetime as dt
import openpyxl
import config
import ezsheets

DRIVER_LOCATION = "/usr/local/bin/chromedriver"
DATA_URL = "https://drive.google.com/file/d/11KF1DuN5tntugNc10ogQDzFnW05ruzLH/view"
XPATH = "/html/body/div[3]/div[3]/div/div[3]/div[2]/div[2]/div[3]"
FILE_NAME = "CityofToronto_COVID-19_Status_Public_Reporting"
FILE_EXTENSIONS = [".xlsx", ".xlsm"]
ACTIVE_ROW = 436
ACTIVE_ROW_LINE_NO = 17
PREVIOUS_DATA = ['2021-05-17', 161904, 150177, 3271, 1011, 271]
PREVIOUS_DATA_LINE_NO = 19
DATE_FORMAT = '%Y-%m-%d'
INDENT = "   "

# # # # # # # #
# DEFINITIONS #
# # # # # # # #

def wait(seconds):
    print(f"{INDENT}Waiting {seconds} seconds...")
    sleep(seconds)

def makeFileNames(fileName, listOfExt, numOfCopies):
    fileNames = []
    for num in range(numOfCopies):
        if num == 0:
            fileNames.append(f"{fileName}{listOfExt[0]}")
            fileNames.append(f"{fileName}{listOfExt[0]}{listOfExt[1]}")
        else:
            fileNames.append(f"{fileName} ({num}){listOfExt[0]}")
            fileNames.append(f"{fileName} ({num}){listOfExt[0]}{listOfExt[1]}")
    return fileNames

def mountDataFile(fileName, folder):
    try:
        print(f"{INDENT}Looking for {fileName} in {folder}...")
        wb = openpyxl.load_workbook(folder + fileName)
        workingPath = folder + fileName
        print(f"{INDENT}Found!")
        return [True, wb, workingPath]
    except:
        print(f"{INDENT}Could not find: {fileName} in {folder}.")
        return [False, None, None]

def updateGoogleSheet(data):
    try:
        print(f"Updating Google Sheet with data: {data}")
        print(f"{INDENT}Accessing Google sheet...")
        doc = ezsheets.Spreadsheet(config.GOOGLE_SHEET_ID)
        sheet = doc[0]
        print(f"{INDENT}Updating row {ACTIVE_ROW}...")
        sheet.updateRow(ACTIVE_ROW, data)
        print(f"{INDENT}Update successful!")
        return data[:6]
    except:
        print(f"{INDENT}Error: Could not update Google Sheet.")
        return False

def updateSourceScript(data):
    # Update counters in in source for next time
    print("Updating source script...")
    print(f"{INDENT}Reading update.py...")
    content = []
    with open(__file__,"r") as f:
        for line in f:
            content.append(line)

    # Write updates to source file
    print(f"{INDENT}Writing update.py...")
    with open(__file__,"w") as f:
        nextRow = ACTIVE_ROW + 1
        print(f"{INDENT}Writing: ACTIVE_ROW = {nextRow}")
        content[ACTIVE_ROW_LINE_NO - 1] = f"ACTIVE_ROW = {nextRow}\n"
        print(f"{INDENT}Writing: PREVIOUS_DATA = {data}")
        content[PREVIOUS_DATA_LINE_NO - 1] = f"PREVIOUS_DATA = {data}\n"
        for i in range(len(content)):
            f.write(content[i])

def isPreviouslyReadDate(sheetDate, prevDate):
    incoming = dt.datetime.strptime(sheetDate, DATE_FORMAT)
    previous = dt.datetime.strptime(prevDate, DATE_FORMAT)
    return incoming <= previous


# # # # # # # # #
# MAIN  PROGRAM #
# # # # # # # # #

def main():

# BROWSER

    # Open browser, download file
    print("Running update.py...")
    print(f"Last date recorded: {PREVIOUS_DATA[0]}")
    print("Getting data file...")
    print(f"{INDENT}Opening browser...")
    browser = webdriver.Chrome(DRIVER_LOCATION)
    browser.get(DATA_URL)
    wait(4)

    # Locate + Click download button
    button = browser.find_element_by_xpath(XPATH)
    print(f"{INDENT}Found <{button.tag_name}> element.")
    button.click()
    print(f"{INDENT}Clicked <{button.tag_name}> element.")
    wait(7)

    # Quit browser
    print(f"{INDENT}Quitting browser...")
    browser.quit()
    wait(4)

# SPREADSHEET

    # Get download file
    print("Searching for download file...")
    result = None
    wb = None
    workingPath = ''
    numOfCopies = 4
    fileNames = makeFileNames(FILE_NAME, FILE_EXTENSIONS, numOfCopies)
    for file in fileNames:
        for folder in config.DOWNLOAD_FOLDER:
            [result, wb, workingPath] = mountDataFile(file, folder)
            if result:
                break
        if result:
            break
    if not wb:
        # Halt if no file was found
        print(f"Could not find file after {numOfCopies} tries.")
        print("Halting with exit code 1")
        return 1

    # Open download file
    print(f"Opening downloaded file...")
    print(f"{INDENT}Retrieving date...")
    dateSheet = wb['Cases by Reported Date']
    SHEET_DATE = dateSheet['A2'].value
    SHEET_DATE = str(SHEET_DATE).split(" ")[0]

    # Halt if we've seen this data before
    if isPreviouslyReadDate(SHEET_DATE, PREVIOUS_DATA[0]):
        print(f"{INDENT}Diverting update: Downloaded file has already been read. (Already read data from: {SHEET_DATE})")
       
        # Copy previous day's data into new row for today
        print(f"{INDENT}Proceeding to update Google Sheet with prev data...")
        newDate = dt.datetime.strptime(PREVIOUS_DATA[0], DATE_FORMAT)
        newDate += dt.timedelta(days=1)
        DUPLICATE_DATA = [newDate.strftime(DATE_FORMAT)] + PREVIOUS_DATA[1:]
        WRITTEN_DATA = updateGoogleSheet(DUPLICATE_DATA)

        if WRITTEN_DATA:
            updateSourceScript(WRITTEN_DATA)
        else:
            print("Error: Could not update source script with prev day's data")
            print("Halting with exit code 3")
            return 3

        # Rename downloaded file for inspection
        print(f"{INDENT}Renaming downloaded file for inspection...")
        newFileName = workingPath + WRITTEN_DATA[0] + FILE_EXTENSIONS[0]
        os_rename(workingPath, newFileName)
        print(f"{INDENT}New file name: {newFileName}")
        print("Halting with exit code 2")
        return 2

    # Get data
    print(f"{INDENT}Retrieving COVID data...")
    sheet = wb['Status']
    TOTAL_CASE_COUNT = sheet['C2'].value
    RECOVERED = sheet['C5'].value
    FATAL = sheet['C6'].value
    CURRENTLY_HOSP = sheet['C8'].value
    CURRENTLY_ICU = sheet['C9'].value

    LATEST_DATA = [SHEET_DATE, TOTAL_CASE_COUNT, RECOVERED, FATAL, CURRENTLY_HOSP, CURRENTLY_ICU]
    print(f"{INDENT}Success! New data retrieved: {LATEST_DATA}")

    # Delete file once we have the data
    print(f"{INDENT}Deleting downloaded file...")
    os_remove(workingPath)

# WRITE TO REMOTE

    WRITTEN_DATA = updateGoogleSheet(LATEST_DATA)

# CLEAN UP

    if WRITTEN_DATA:
        updateSourceScript(WRITTEN_DATA)
    else:
        print("Error: Could not update source script with today's data")
        print("Halting with exit code 4")
        return 4

    print("Success! All done!")
    return 0

if __name__ == "__main__":
    main()
