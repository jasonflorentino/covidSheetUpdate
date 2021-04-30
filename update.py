#! Python3

from selenium import webdriver
from time import sleep
from os import remove as os_remove
import openpyxl
import config
import ezsheets

DRIVER_LOCATION = "/usr/local/bin/chromedriver"
DATA_URL = "https://drive.google.com/file/d/11KF1DuN5tntugNc10ogQDzFnW05ruzLH/view"
XPATH = "/html/body/div[3]/div[3]/div/div[3]/div[2]/div[2]/div[3]"
FILE_NAME = "CityofToronto_COVID-19_Status_Public_Reporting"
FILE_EXTENSIONS = [".xlsx", ".xlsm"]
ACTIVE_ROW = 418
ACTIVE_ROW_LINE_NO = 15
PREVIOUS_DATE = '2021-04-29'
PREVIOUS_DATE_LINE_NO = 17
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

def mountFile(fileName, folder):
    try:
        print(f"{INDENT}Looking for {fileName} in {folder}...")
        wb = openpyxl.load_workbook(folder + fileName)
        workingPath = folder + fileName
        print(f"{INDENT}Found!")
        return [True, wb, workingPath]
    except:
        print(f"{INDENT}Could not find: {fileName} in {folder}.")
        return [False, None, None]

# # # # # # # # #
# MAIN  PROGRAM #
# # # # # # # # #

def main():

# BROWSER

    # Open browser, download file
    print("Running update.py...")
    print(f"Last date recorded: {PREVIOUS_DATE}")
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
            [result, wb, workingPath] = mountFile(file, folder)
            if result:
                break
        if result:
            break
    if not wb:
        # Halt if no file was found
        print(f"Could not find file after {numOfCopies} tries.")
        return 1

    # Open download file
    print(f"Opening downloaded file...")
    print(f"{INDENT}Retrieving date...")
    dateSheet = wb['Cases by Reported Date']
    DATE = dateSheet['A2'].value
    DATE = str(DATE).split(" ")[0]

    # Halt if we've seen this data before
    if DATE == PREVIOUS_DATE:
        print(f"{INDENT}Stopping execution: Spreadsheet data has already been read. (Already read data from: {DATE})")
        print(f"{INDENT}Deleting downloaded file...")
        os_remove(workingPath)
        return 2

    # Get data
    print(f"{INDENT}Retrieving COVID data...")
    sheet = wb['Status']
    TOTAL_CASE_COUNT = sheet['B2'].value
    RECOVERED = sheet['B5'].value
    FATAL = sheet['B6'].value
    CURRENTLY_HOSP = sheet['B8'].value
    CURRENTLY_ICU = sheet['B9'].value

    LATEST_DATA = [DATE, TOTAL_CASE_COUNT, RECOVERED, FATAL, CURRENTLY_HOSP, CURRENTLY_ICU]
    print(f"{INDENT}Success! New data retrieved: {LATEST_DATA}")

    # Delete file once we have the data
    print(f"{INDENT}Deleting downloaded file...")
    os_remove(workingPath)

# WRITE TO REMOTE

    # Update Google Sheet
    print("Updating Google Sheet with new data...")
    print(f"{INDENT}Accessing Google sheet...")
    doc = ezsheets.Spreadsheet(config.GOOGLE_SHEET_ID)
    sheet = doc[0]
    print(f"{INDENT}Updating row {ACTIVE_ROW}...")
    sheet.updateRow(ACTIVE_ROW, LATEST_DATA)

# CLEAN UP

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
        print(f"{INDENT}Writing: PREVIOUS_DATE = '{DATE}'")
        content[PREVIOUS_DATE_LINE_NO - 1] = f"PREVIOUS_DATE = '{DATE}'\n"
        for i in range(len(content)):
            f.write(content[i])

    print("Success! All done!")
    return 0

if __name__ == "__main__":
    main()
