#! Python3

"""
Imports
"""
import requests
import openpyxl
import config
import datetime as dt
import ezsheets
import os

"""
Script Constants
"""
DOWNLOAD_URL = "https://drive.google.com/u/0/uc?id=11KF1DuN5tntugNc10ogQDzFnW05ruzLH&export=download"
FILE_NAME = "toronto_covid_data.xlsx"
FULL_PATH_TO_FILE = config.DESKTOP + FILE_NAME
DATE_FORMAT = '%Y-%m-%d'
DATE_FORMAT_LONG = '%B %d , %Y'
INDENT = "   "

"""
Sheet Update Values
"""
ACTIVE_ROW = 545
ACTIVE_ROW_LINE_NO = 26
PREVIOUS_DATA = ['2021-09-03', 174911, 169644, 3622, 75, 22]
PREVIOUS_DATA_LINE_NO = 28
DAYS_SINCE_NO_UPDATE = 0
DAYS_SINCE_NO_UPDATE_LINE_NO = 30

"""
Definitions
"""
def log(msg, level=1):
  space = INDENT * (level - 1)
  print(space + msg)

def downloadFile(url, outputFile):
  res = requests.get(url, allow_redirects=True, stream=True)
  with open(outputFile, 'wb') as file:
    for chunk in res.iter_content():
      file.write(chunk)
  file.close()

def isPreviouslyReadDate(newDate, prevDate):
  incoming = dt.datetime.strptime(newDate, DATE_FORMAT)
  previous = dt.datetime.strptime(prevDate, DATE_FORMAT)
  return incoming <= previous

def updateGoogleSheet(data):
  try:
    log(f"Updating Google Sheet with data: {data}")
    log("Accessing Google sheet...", 2)
    doc = ezsheets.Spreadsheet(config.GOOGLE_SHEET_ID)
    sheet = doc[0]
    log(f"Updating row {ACTIVE_ROW}...", 2)
    sheet.updateRow(ACTIVE_ROW, data)
    log("Update successful!", 2)
    return data[:6]
  except:
    log("Error: Could not update Google Sheet.", 2)
    return False

def updateSourceScript(data):
    # Update counters in in source for next time
    log("Updating source script...")
    log("Reading update.py...", 2)
    content = []
    with open(__file__,"r") as f:
        for line in f:
            content.append(line)

    # Write updates to source file
    log("Writing update.py...", 2)
    with open(__file__,"w") as f:
        nextRow = ACTIVE_ROW + 1
        log(f"Writing: ACTIVE_ROW = {nextRow}", 2)
        content[ACTIVE_ROW_LINE_NO - 1] = f"ACTIVE_ROW = {nextRow}\n"
        log(f"Writing: PREVIOUS_DATA = {data}", 2)
        content[PREVIOUS_DATA_LINE_NO - 1] = f"PREVIOUS_DATA = {data}\n"
        log(f"Writing: DAYS_SINCE_NO_UPDATE = {DAYS_SINCE_NO_UPDATE}", 2)
        content[DAYS_SINCE_NO_UPDATE_LINE_NO - 1] = f"DAYS_SINCE_NO_UPDATE = {DAYS_SINCE_NO_UPDATE}\n"
        for i in range(len(content)):
            f.write(content[i])

"""
Main Program
"""
def main():
  global DAYS_SINCE_NO_UPDATE

  log("Downloading file...")
  downloadFile(DOWNLOAD_URL, FULL_PATH_TO_FILE)

  # Open download file
  log("Opening downloaded file...")
  wb = openpyxl.load_workbook(FULL_PATH_TO_FILE)
  log("Last date recorded: " + PREVIOUS_DATA[0], 2)
  log("Retrieving spreadsheet's date...", 2)
  dateSheet = wb['Data Note']
  textDate = str(dateSheet['A2'].value)[11:]
  dtInstance = dt.datetime.strptime(textDate, DATE_FORMAT_LONG)
  sheetDate = dt.datetime.strftime(dtInstance, DATE_FORMAT)
  log("Incoming date: " + sheetDate, 2)

  if isPreviouslyReadDate(sheetDate, PREVIOUS_DATA[0]):
    log(f"Diverting update: Downloaded file has already been read. (Already read data from: {sheetDate})", 2)
       
    # Copy previous day's data into new row for today
    log("Proceeding to update Google Sheet with prev data...", 2)
    newDate = dt.datetime.strptime(PREVIOUS_DATA[0], DATE_FORMAT)
    DAYS_SINCE_NO_UPDATE += 1
    newDate += dt.timedelta(days=DAYS_SINCE_NO_UPDATE)
    DUPLICATE_DATA = [newDate.strftime(DATE_FORMAT)] + PREVIOUS_DATA[1:]
    WRITTEN_DATA = updateGoogleSheet(DUPLICATE_DATA)

    if WRITTEN_DATA:
        updateSourceScript(PREVIOUS_DATA)
    else:
        log("Error: Could not update source script with prev day's data")
        log("Halting with exit code 3")
        return 3

    # Rename downloaded file for inspection
    log("Renaming downloaded file for inspection...", 2)
    newFileName = FULL_PATH_TO_FILE + WRITTEN_DATA[0] + "xlsx"
    os.rename(FULL_PATH_TO_FILE, newFileName)
    log("New file name: " + newFileName, 2)
    log("Halting with exit code 2")
    return 2

  # Get data
  log("Retrieving COVID data...", 2)
  sheet = wb['Status']
  totalCaseCount = sheet['C2'].value
  recoveredCases = sheet['C5'].value
  fatalCases = sheet['C6'].value
  currentlyHospitalized = sheet['C8'].value
  currentlyInICU = sheet['C9'].value

  latestData = [
    sheetDate, 
    totalCaseCount, 
    recoveredCases, 
    fatalCases, 
    currentlyHospitalized, 
    currentlyInICU
  ]

  log(f"Success! New data retrieved: {latestData}", 2)

  # Delete file once we have the data
  log("Deleting downloaded file...", 2)
  os.remove(FULL_PATH_TO_FILE)

# WRITE TO REMOTE

  WRITTEN_DATA = updateGoogleSheet(latestData)

# CLEAN UP

  if WRITTEN_DATA:
      DAYS_SINCE_NO_UPDATE = 0
      updateSourceScript(WRITTEN_DATA)
  else:
      log("Error: Could not update source script with today's data")
      log("Halting with exit code 4")
      return 4

  log("Success! All done!")
  return 0

if __name__ == "__main__":
  main()