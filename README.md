# sheets-api
## Description
Python sheets is an independent project aimed at saving time and increasing efficiency by automating repeated entries on google spreadsheets using Google’s **Sheets API**. The script for this project was written for a spreadsheet that was managing my work hours and pay on Python 3. The script currently also utilises the **gspreads** package, which can do simple reads and writes, but cannot make complex changes like updating multiple cells or changing sheet properties. I intend to change the methods that utilises **gspreads** as it has demonstrated some inefficiencies in terms of processing time and convert them to use Google’s **Sheets API**.

## How to use
As it is currently set up, the user interacts with the programme via console input and output. When the programme starts, the console will ask “What do you want to do?” There are four commands that the user can choose from: *Total Earnings, Unpaid Hours, Update or Quit*. *Total earnings* will return the overall earnings from the job, *unpaid hours* will return the number of work hours that was not paid for yet, *update* will allow the user to update (i.e. create a new work hour entry), which will ask the users a series of question to determine the information to update the spreadsheet. 

## OAuth and Key
As with all APIs, OAuth is a key part of making the script work. The repository does not include the credentials keys, as it needs to remain confidential and shouldn’t be published online. 

## Comments
This project is still a work in progess. I am working to implement a Graphical User Interface (GUI) to enhance the user experience. Furthermore, althoug this script only works on my spreadsheet, I believe that it can easily be modified to fit the needs of another spreadsheet. Since there is not many guides to using **Sheets API** online (and this really isn't one), but I hope that this can give a vague idea of how it can be used for other users as well.
