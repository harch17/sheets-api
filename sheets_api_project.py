import datetime as dt
from datetime import *
import sys
import re
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from googleapiclient import discovery


scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
service = discovery.build('sheets', 'v4', credentials=creds)

# Spreadsheet ID
spreadsheetId = '1077v4MI09ZDjdiH8ikNJg55nN0qHAXJiqNI6QzIxEgk'

dateformat = '"%m/%d/%Y %H:%M:%S"'

with open('student_list', 'r') as f:
    s = f.read().strip()

std_lst = re.split(r'\s', s)


def getEarnings(earningStr, total=False):  # Total is false by default
    earningStr = earningStr.lower()

    if earningStr in std_lst:
        range_ = earningStr + '!'
    else:
        print("No student named {0} !".format(earningStr))
        return None  # exits method without returning anything

    # The A1 notation of the values to retrieve.
    range_ = range_ + convertToA1(1, 8)

    if total:
        valRenderOption = 'UNFORMATTED_VALUE'
    else:
        valRenderOption = 'FORMATTED_VALUE'

    request = service.spreadsheets().values().\
        get(spreadsheetId=spreadsheetId,
            range=range_,
            valueRenderOption=valRenderOption)

    response = request.execute()

    return(response['values'][0][0])


def getTotalEarnings():
    rawEarning = getEarnings('minju', True) + getEarnings('minsuk', True)
    print("₩{:,}".format(rawEarning))


# need to change
def calcUnpaidHrs(unpaidStr):
    unpaidStr = unpaidStr.lower()

    if unpaidStr == 'minju':
        range_ = 'Minju!'
    elif unpaidStr == 'minsuk':
        range_ = 'Minsuk!'
    else:
        print("No student named " + unpaidStr + "!")
        return None  # exits method without returning anything

    range_ = range_ + 'D2:E'

    valRenderOption = 'UNFORMATTED_VALUE'

    request = service.spreadsheets().values().\
        get(spreadsheetId=spreadsheetId,
            range=range_,
            valueRenderOption=valRenderOption)

    response = request.execute()

    values = response['values']

    unpaidDays = 0  # Initialise variable

    # Add up all unpaid hours
    for i in values:
        if i[1] == 'N':
            unpaidDays += i[0]

    #   Change time from days to hours in string form
    unpaidHrs = str(dt.timedelta(days=unpaidDays))

    return unpaidHrs


def convertToA1(rows, col):
    aASCII = ord('A')  # Int value of 'A'
    colASCII = aASCII + (col - 1)
    formatedCell = chr(colASCII) + str(rows)
    return formatedCell


# Delete function if not necessary
def convertTimeToSec(n):
    (h, m) = n.split(':')
    secondResult = int(h) * 3600 + int(m) * 60
    return secondResult


def updateLesson():
    date_error_str = "Invalid format for date! Format: mm/dd/yyyy"
    month_error_str = "Month should be between 1 and 12!"
    day_error_str = "Day should be between 1 and 31!"
    year_error_str = "Year is probably 2019"
    date_correct_str = "Format: Format: mm/dd/yyyy"

    lessonDate = ''
    startTime = ''
    endTime = ''

    while True:
        dateCommand = input("Input date of lesson. If today, type today: ").\
            lower()

        if dateCommand == 'today':
            lessonDate = str(date.today())
            (y, m, d) = lessonDate.split('-')
            lessonDate = str(m) + '/' + str(d) + '/' + str(y)
        elif re.match(r'\d?\d\/\d{2}\/\d{4}', dateCommand):
            date_split = re.split(r'/', dateCommand)
            if int(date_split[0]) not in range(1, 13):
                #print(month_error_str)
                print(type(date_split[0]))
                continue
            elif int(date_split[1]) not in range(1, 32):
                print(day_error_str)
                continue
            elif int(date_split[2]) != 2019:
                print(year_error_str)
                continue
            else:
                lessonDate = dateCommand
        else:
            print(date_error_str)
            print(date_correct_str)
            continue

        print("Confirming date:", lessonDate)
        confirm = input("Is this correct? (Y/N)? ").lower()

        if confirm == 'y':
            break
        else:
            continue

    time_error_str = "Invalid format for Time! Format: hh:mm AMorPM"

    while True:
        timeCommand = input(
            "Input the time the lesson started (i.e. 11:00 AM): ").lower()

        if not re.match(r'\d?\d:\d{2}\s(am|pm)', timeCommand):
            print(time_error_str)
            continue

        time_split = re.split(r':|\s', timeCommand)
        h, m = int(time_split[0]), int(time_split[1])

        # Change to 24 hour time
        if time_split[2] == 'pm' and h != 12:
            h += 12

        try:
            startTime = dt.time(h, m)
        except ValueError:
            print("Invalid range for hours and/or minutes")
            continue

        print("Confirming start time:", startTime)
        confirm = input("Is this correct? (Y/N)? ").lower()

        if confirm == 'y':
            break
        else:
            continue

    while True:
        timeCommand = input(
            "Input the time the lesson started (i.e. 11:00 AM): ").lower()

        if not re.match(r'\d?\d:\d{2}\s(am|pm)', timeCommand):
            print(time_error_str)
            continue

        time_split = re.split(r':|\s', timeCommand)
        h, m = int(time_split[0]), int(time_split[1])

        # Change to 24 hour time
        if time_split[2] == 'pm' and h != 12:
            h += 12

        try:
            endTime = dt.time(h, m)
        except ValueError:
            print("Invalid range for hours and/or minutes")
            continue

        print("Confirming start time:", endTime)
        confirm = input("Is this correct? (Y/N)? ").lower()

        if confirm == 'y':
            break
        else:
            continue

    while True:
        student = input("Which student? ").lower()

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

    pushUpdate(student_sheet, lessonDate, startTime, endTime, duration)


def pushUpdate(sheet, date, time_1, time_2, dur):
    batch_update_spreadsheet_request_body = {
        # A list of updates to apply to the spreadsheet.
        # Requests will be applied in the order they are specified.
        # If any request is not valid, no requests will be applied.
        "requests": [
            {  # Insert row
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
            {  # Bold First Row
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
            {  # Time format column B, C
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
            {  # Validation for column E
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
            {  # Enter Data
                "pasteData": {
                    "data": concat([date, time_1, time_2, dur, 'N']),
                    "type": "PASTE_NORMAL",
                    "delimiter": ",",
                    "coordinate": {
                        "sheetId": sheet,  # 1st sheet
                        "rowIndex": 1  # 2nd row
                    }
                }
            }
        ]
    }
    request = service.spreadsheets().\
        batchUpdate(spreadsheetId=spreadsheetId,
                    body=batch_update_spreadsheet_request_body)

    request.execute()


def concat(lst):
    if len(lst) == 0:
        return ''
    elif len(lst) == 1:
        return str(lst[0])
    else:
        return str(lst[0]) + ', ' + concat(lst[1:])


# def updateWork()


def main():
    while(True):
        command = ''
        command = input("What do you want to do? ").lower()

        if command == 'help':
            print("Here are the list of commands you can use: \
                  \n - Total Earnings \n - Earnings \n - Unpaid Hours \
                  \n - New Entry \n - Quit")

        if command == 'total earnings':
            getTotalEarnings()

        if command == 'unpaid hours':
            while(True):
                unpaidCmd = input("Which Student? (Minju or Minsuk) ")

                unpaid = calcUnpaidHrs(unpaidCmd)
                if unpaid is None:
                    continue
                else:
                    print(unpaid)
                    break

        if command == 'new entry':
            updateLesson()

        if command == 'quit':
            print()
            sys.exit(0)

        if command == 'earnings':
            while(True):
                earnCmd = input("Which Student? (Minju or Minsuk) ")

                earn = getEarnings(earnCmd)
                if earn is None:
                    continue
                else:
                    print(earn)
                    break


main()
