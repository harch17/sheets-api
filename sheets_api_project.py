import gspread as gs
import datetime as dt
from datetime import *
import sys
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
#from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import discovery


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
gc = gs.authorize(creds)
service = discovery.build('sheets', 'v4', credentials=creds)
 
# Open spreadsheet
wks = gc.open("Work + Tutoring").sheet1

# Spreadsheet ID
spreadsheetId = '1077v4MI09ZDjdiH8ikNJg55nN0qHAXJiqNI6QzIxEgk'

dateformat = '"%m/%d/%Y %H:%M:%S"'

def getEarnings(earningStr):
	getSheet = 0 # default sheet is first sheet

	if earningStr == 'minju':
		getSheet = 0
	elif earningStr == 'minsuk':
		getSheet = 1741513205
	else:
		print("No student named " + earningStr + "!")
		return None # exits method without returning anything

	# The ID of the spreadsheet to retrieve data from.
	spreadsheet_id = getSheet 

	# The A1 notation of the values to retrieve.
	range_ = convertToA1(1, 8) + ':' + convertToA1(1, 8)

	# How values should be represented in the output.
	# The default render option is ValueRenderOption.FORMATTED_VALUE.
	value_render_option = 'UNFORMATTED_VALUE'

	request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_, valueRenderOption=value_render_option)
	response = request.execute()

	print(response)

# need to change
def getTotalEarnings():
	getEarnings('minju') + getEarnings('minsuk')




	getRequest = service.spreadsheets().getByDataFilter(spreadsheetId=spreadsheet_id, body=get_spreadsheet_by_data_filter_request_body)
	response = request.execute()
	# Position of Earnings Cell
	earningsRow = wks.find("Total Earnings").row 
	earningsCol = wks.find("Total Earnings").col
	print(wks.cell(earningsRow, earningsCol + 1).value)

# need to change
def calcUnpaidHrs():
	unpaidSec = 0
	# Column containing whether I was paid or not
	# payCol = wks.find("Paid").col
	notPaidCells = wks.findall("N")
	# Non-paid Cell = NPC, containing how long each session
	firstNPC = convertToA1(notPaidCells[0].row, notPaidCells[0].col - 1)
	lastNPC = convertToA1(notPaidCells[len(notPaidCells) - 1].row, notPaidCells[len(notPaidCells) - 1].col - 1)

	NPCRange = firstNPC + ':' + lastNPC

	timeFetch = wks.range(NPCRange)

	# Calculate unpaid time in terms of seconds
	for time in timeFetch:
		unpaidSec += convertTimeToSec(time.value)

	# Change time from seconds to hours in string form
	unpaidHrs = str(dt.timedelta(seconds=unpaidSec))
	print(unpaidHrs)


def convertToA1(rows, col):
	# Int value of 'A'
	aASCII = ord('A') 
	colASCII = aASCII + (col - 1)
	formatedCell = chr(colASCII) + str(rows)
	return formatedCell


def convertTimeToSec(n):
	(h, m) = n.split(':')
	secondResult = int(h) * 3600 + int(m) * 60
	return secondResult


def update():
	lessonDate = ''
	startTime = ''
	endTime = ''

	# Determine Lesson Date
	while(True):
		dateCommand = ''
		try:
			dateCommand = input("Input date of lesson. If today, type today: ")
		except EOFError: # EOF command quits programme
			print()
			sys.exit(0)

		dateCommand = dateCommand.lower()

		if dateCommand == 'today':
			lessonDate = str(date.today())
			(y, m, d) = lessonDate.split('-')
			lessonDate = str(m) + '/' + str(d) + '/' + str(y)

			print(lessonDate)

			break
		else:
			lessonDate = dateCommand
			try:
				(m, d, y) = lessonDate.split('/')
			except ValueError:
				print("Invalid format for date! Format: mm/dd/yy")
				continue

			(m, d, y) = lessonDate.split('/')

			if int(m) > 12 or int(m) < 1:
				print("Month should be between 1 and 12!")
				print("Format: Format: mm/dd/yy")
				continue

			if int(d) > 31 or int(d) < 1:
				print("Day should be between 1 and 31!")
				print("Format: Format: mm/dd/yy")
				continue

			if int(y) != 2019:
				print("Year is probably 2019")
				continue

			print("Confirming date:", lessonDate)
			confirm = input("Is this correct? (Y/N)? ").lower()

			if confirm == 'y':
				break
			else:
				continue

	# Determine start and end time
	while(True):
		timeCommand = ''
		try:
			timeCommand = input("Input the time the lesson started (i.e. 11:00 AM): ")
		except EOFError: # EOF command quits programme
			print()
			sys.exit(0)

		timeCommand = timeCommand.lower()

		try:
			(h, minAndAMPM) = timeCommand.split(':')
		except ValueError:
			print("Invalid format for Time! Format: hh:mm AMorPM 1")
			continue

		(h, minAndAMPM) = timeCommand.split(':')

		try:
			(m, AMPM) = minAndAMPM.split(' ')
		except ValueError:
			print("Invalid format for Time! Format: hh:mm AMorPM 2")
			continue

		(m, AMPM) = minAndAMPM.split(' ')

		if AMPM != 'am' and AMPM != 'pm':
			print("Invalid format for Time! Format: hh:mm AMorPM 3")
			continue

		if AMPM == 'pm' and h != 12:
			h = h + 12

		h = int(h)
		m = int(m)

		try:
			startTime = dt.time(h, m)
		except ValueError:
			print("Invalid range for hours and/or minutes")
			continue

		startTime = dt.time(h, m)

		print("Confirming start time:", startTime)
		confirm = input("Is this correct? (Y/N)? ").lower()

		if confirm == 'y':
			break
		else:
			continue

	while(True):
		timeCommand = ''
		try:
			timeCommand = input("Input the time the lesson ended (i.e. 11:00 AM): ")
		except EOFError: # EOF command quits programme
			print()
			sys.exit(0)

		timeCommand = timeCommand.lower()

		try:
			(h, minAndAMPM) = timeCommand.split(':')
		except ValueError:
			print("Invalid format for Time! Format: hh:mm AMorPM")
			continue

		(h, minAndAMPM) = timeCommand.split(':')

		try:
			(m, AMPM) = minAndAMPM.split(' ')
		except ValueError:
			print("Invalid format for Time! Format: hh:mm AMorPM")
			continue

		(m, AMPM) = minAndAMPM.split(' ')

		if AMPM != 'am' and AMPM != 'pm':
			print("Invalid format for Time! Format: hh:mm AMorPM")
			continue

		h = int(h)
		m = int(m)

		if AMPM == 'pm' and h != 12:
			h = h + 12

		try:
			endTime = dt.time(h, m)
		except ValueError:
			print("Invalid range for hours and/or minutes")
			continue

		endTime = dt.time(h, m)

		print("Confirming end time:", endTime)
		confirm = input("Is this correct? (Y/N)? ").lower()

		if confirm == 'y':
			break
		else:
			continue

	while(True):
		student = ''
		try:
			student = input("Which student? ").lower()
		except EOFError: # EOF command quits programme
			print()
			sys.exit(0)

		if student == 'minju':
			student_sheet = 0
			break
		elif student == 'minsuk':
			student_sheet = 1741513205
			break
		else:
			print("That student doesn't exist!")
			continue

	time1 = dt.datetime.combine(dt.date.today(), startTime)
	time2 = dt.datetime.combine(dt.date.today(), endTime)
	duration = str(dt.timedelta(seconds=(time2 - time1).total_seconds()))[:-3]

	pushUpdate(student_sheet,lessonDate, startTime, endTime, duration)


def pushUpdate(sheet, date, time_1, time_2, dur):
	batch_update_spreadsheet_request_body = {
    # A list of updates to apply to the spreadsheet.
    # Requests will be applied in the order they are specified.
    # If any request is not valid, no requests will be applied.
    "requests": [
    	{ # Insert row
      	"insertDimension": {
        	"range": {
          	"sheetId": sheet,
          	"dimension": "ROWS",
          	"startIndex": 1,
          	"endIndex": 2
        	},
        	"inheritFromBefore": False
      	}
    	},
    	{ # Bold First Row
    		"repeatCell": {
    			"range": {
    				"sheetId": sheet,
    				"endRowIndex": 1
    			},
    			"cell": {
    				"userEnteredFormat": {
    					"textFormat": {
    						"bold": True
    					}
    				}
    			},
    			"fields": "userEnteredFormat.textFormat.bold"
    		}
    	},
    	{ # Time format column B, C
    		"repeatCell": {
    			"range": {
    				"sheetId": sheet,
    				"startRowIndex": 1,
    				"startColumnIndex": 1,
    				"endColumnIndex": 3
    			},
    			"cell": {
    				"userEnteredFormat": {
    					"numberFormat": {
    						"type": "DATE",
    						"pattern": 'hh:mm A/P".M."'
    					}
    				}
    			},
    			"fields": "userEnteredFormat.numberFormat"
    		}
    	},
    	{ # Validation for column E
    		"setDataValidation": {
    			"range": {
    				"sheetId": sheet,
    				"startRowIndex": 1,
    				"startColumnIndex": 4,
    				"endColumnIndex": 5
    			},
    			"rule": {
    				"condition": {
    					"type": "ONE_OF_LIST",
    					"values": [
    						{"userEnteredValue": "Y"},
    						{"userEnteredValue": "N"}
    					]
    				}
    			}
    		}
    	},
    	{ # Enter Data
    		"pasteData": {
    			"data": concat([date, time_1, time_2, dur, 'N']),
    			"type": "PASTE_NORMAL",
    			"delimiter": ",",
    			"coordinate": {
    				"sheetId": sheet, # 1st sheet
    				"rowIndex": 1 # 2nd row
    			}
    		}
	    }
  	]
  }
	request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=batch_update_spreadsheet_request_body)
	response = request.execute()


# Input list and will create one string that contains all items separated by commas
def concat(list_in):
	string = ''
	for obj in list_in[:-1]:
		string += str(obj) + ', '
	return string + str(list_in[-1])


def main():
	while(True):
		command = ''
		try:
			command = input("What do you want to do? ")
		except EOFError: # EOF command quits programme
			print()
			sys.exit(0)

		command = command.lower()

		if command == 'help':
			print("Here are the list of commands you can use: \n - Total Earnings \n - Unpaid Hours \n - Update \n - Quit")

		if command == 'total earnings':
			getTotalEarnings()

		if command == 'unpaid hours':
			calcUnpaidHrs()

		if command == 'update':
			update()

		if command == 'quit':
			print()
			sys.exit(0)

		if command == 'test':
			getEarnings('minju')

main()
