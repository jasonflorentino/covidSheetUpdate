#! Python3
"""
V2 of update.py downloads the data file directly
instead of using selenium to navigate to the webpage 
to download the file.
— Jason, 2021-09-05
"""

"""
Imports
"""
import requests
import openpyxl
import config
import datetime as dt
import ezsheets
import os
import sys

"""
Script Constants
"""
DOWNLOAD_URL = "https://drive.google.com/u/0/uc?id=11KF1DuN5tntugNc10ogQDzFnW05ruzLH&export=download"
FILE_NAME = "toronto_covid_data.xlsx"
FULL_PATH_TO_FILE = config.DESKTOP + FILE_NAME
DATE_FORMAT = '%Y-%m-%d'
DATE_FORMAT_LONG = '%B %d, %Y'
INDENT = "   "

"""
Sheet Update Values
"""
ACTIVE_ROW = 549
ACTIVE_ROW_LINE_NO = 33
PREVIOUS_DATA = ['2021-09-07', 175496, 170229, 3624, 83, 24]
PREVIOUS_DATA_LINE_NO = 35
DAYS_SINCE_NO_UPDATE = 0
DAYS_SINCE_NO_UPDATE_LINE_NO = 37

# # # # # # # # # #
# Fn  Definitions #
# # # # # # # # # #

"""
Logs message to std out
msg   - Message to log
level - Indent level to print at (default 1, no indent) 
"""
def log(msg, level=1):
	space = INDENT * (level - 1)
	print(space + msg)

"""
Downloads file from given url, to a given outputFile
url        - Url to download, assumes chunked content
outputFile - Destination file
"""
def downloadFile(url, outputFile):
	try:
		res = requests.get(url, allow_redirects=True, stream=True)
		with open(outputFile, 'wb') as file:
			for chunk in res.iter_content():
				file.write(chunk)
		file.close()
	except requests.exceptions.RequestException as e:
		log("File download failed!")
		raise SystemExit(e)

"""
Returns true if the newDate is earlier or the same as the prevDate
"""
def isPreviouslyReadDate(newDate, prevDate):
	incoming = dt.datetime.strptime(newDate, DATE_FORMAT)
	previous = dt.datetime.strptime(prevDate, DATE_FORMAT)
	return incoming <= previous

"""
Writes given data to the ACTIVE_ROW in the Google Sheet
"""
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

"""
Rewrites this file with updates for next execution
"""
def updateSourceScript(data):
	log("Updating source script...")
	log("Reading update.py...", 2)
	# Read file contents
	content = []
	with open(__file__,"r") as f:
		for line in f:
			content.append(line)
	# Write updates to source file
	log("Writing update.py...", 2)
	with open(__file__,"w") as f:
		log(f"Writing: ACTIVE_ROW = {ACTIVE_ROW}", 2)
		log(f"Writing: PREVIOUS_DATA = {data}", 2)
		log(f"Writing: DAYS_SINCE_NO_UPDATE = {DAYS_SINCE_NO_UPDATE}", 2)
		content[ACTIVE_ROW_LINE_NO - 1] = f"ACTIVE_ROW = {ACTIVE_ROW}\n"
		content[PREVIOUS_DATA_LINE_NO - 1] = f"PREVIOUS_DATA = {data}\n"
		content[DAYS_SINCE_NO_UPDATE_LINE_NO - 1] = f"DAYS_SINCE_NO_UPDATE = {DAYS_SINCE_NO_UPDATE}\n"
		for i in range(len(content)):
			f.write(content[i])

"""
Returns given workbook's date as string in format YYYY-MM-DD
wb - Work book to pull date out of
"""
def getSheetDate(wb):
	log("Retrieving spreadsheet's date...", 2)
	cellText = str(wb['Data Note']['A2'].value)
	textDate = cellText[11:] # Slice date from string: 'Data as of Month DD , YYYY'
	dtInstance = dt.datetime.strptime(textDate, DATE_FORMAT_LONG)
	return dt.datetime.strftime(dtInstance, DATE_FORMAT)

"""
Pulls covid data out of a given sheet and returns as a list with the given date
"""
def getCovidData(sheet, date):
	totalCaseCount = sheet['C2'].value
	recoveredCases = sheet['C5'].value
	fatalCases = sheet['C6'].value
	currentlyHospitalized = sheet['C8'].value
	currentlyInICU = sheet['C9'].value
	return [
		date, 
		totalCaseCount, 
		recoveredCases, 
		fatalCases, 
		currentlyHospitalized, 
		currentlyInICU
	]

# # # # # # # # #
# Main  Program #
# # # # # # # # #

def main():
	global DAYS_SINCE_NO_UPDATE
	global ACTIVE_ROW
	log("=== Running update-v2.py ===")

	log("Downloading file...")
	downloadFile(DOWNLOAD_URL, FULL_PATH_TO_FILE)

	log("Opening downloaded file...")
	wb = openpyxl.load_workbook(FULL_PATH_TO_FILE)

	log("Comparing dates...")
	sheetDate = getSheetDate(wb)
	log("Last date recorded: " + PREVIOUS_DATA[0], 2)
	log("Incoming date: " + sheetDate, 2)

	if isPreviouslyReadDate(sheetDate, PREVIOUS_DATA[0]):
		log(f"Error: Downloaded file has already been read. Already read data from: {sheetDate}", 2)

		# Increment date and combine with old data
		newDate = dt.datetime.strptime(PREVIOUS_DATA[0], DATE_FORMAT)
		DAYS_SINCE_NO_UPDATE += 1
		newDate += dt.timedelta(days=DAYS_SINCE_NO_UPDATE)
		DUPLICATE_DATA = [newDate.strftime(DATE_FORMAT)] + PREVIOUS_DATA[1:]
		
		log("Proceeding to update Google Sheet with prev data...", 2)
		WRITTEN_DATA = updateGoogleSheet(DUPLICATE_DATA)

		if WRITTEN_DATA:
			ACTIVE_ROW += 1
			updateSourceScript(PREVIOUS_DATA)
		else:
			log("Error: Google Sheet update failed. Script rewrite was not executed.")
			return "Halting with exit code 1"

		log("Renaming downloaded file for inspection...", 2)
		newFileName = FULL_PATH_TO_FILE + WRITTEN_DATA[0] + ".xlsx"
		
		log("New file name: " + newFileName, 2)
		os.rename(FULL_PATH_TO_FILE, newFileName)

		return "Halting with exit code 2"

	# Date is new, continue as normal
	log("New date!")
	log("Retrieving COVID data...", 2)
	sheet = wb['Status']
	latestData = getCovidData(sheet, sheetDate)
	log(f"New data retrieved: {latestData}", 2)

	WRITTEN_DATA = updateGoogleSheet(latestData)

	if WRITTEN_DATA:
		ACTIVE_ROW += 1
		DAYS_SINCE_NO_UPDATE = 0
		updateSourceScript(WRITTEN_DATA)
	else:
		log("Error: Could not update source script with today's data")
		return "Halting with exit code 3"

	log("Success! All done!")
	return 0

if __name__ == "__main__":
	sys.exit(main())