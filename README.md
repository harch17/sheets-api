# sheets-api
## Description
Python sheets is an independent project aimed at saving time and increasing efficiency by automating repeated entries on google spreadsheets using Google’s **Sheets API**. The script for this project was written for a spreadsheet that was managing my work hours and pay on Python 3. 

## How to use
As it is currently set up, the user interacts with the programme via console input and output. When the programme starts, the console will ask “What do you want to do?” There are four commands that the user can choose from: *Total Earnings, Unpaid Hours, Update, Earnings or Quit*. *Total earnings* will return the overall earnings, *unpaid hours* will return the number of work hours that was not paid for yet, *update* will allow the user to update (i.e. create a new work hour entry), which will ask the users a series of question to determine the information to update the spreadsheet and *earnings* will allow the user to specify which earnings to return.

## OAuth and Key
As with all APIs, OAuth is a key part of making the script work. The repository does not include the credentials keys, as it needs to remain confidential and shouldn’t be published online. 

## Comments
This project is still a work in progess. I am working to implement a Graphical User Interface (GUI) to enhance the user experience. Furthermore, althoug this script only works on my spreadsheet, I believe that it can easily be modified to fit the needs of another spreadsheet. Since there is not many guides to using **Sheets API** online (and this really isn't one), but I hope that this can give a vague idea of how it can be used for other users as well.
