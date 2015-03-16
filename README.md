# Send-Mails-From-Marks-Inside-a-Spreadsheet

The code Send Mails containging Marks in a Google Spreadhsheet formatted like this:
https://docs.google.com/spreadsheets/d/1JobjeymNWivpwTH1x2syL9pUill9kilF8Bx1ro19lnA/edit#gid=0

It's a quick project to help students receiving detailled marks ont their Practical lessons.
  => If anyone is intersting by this. I can make it easier to use.
Performances are not the aim of the project, I prefered readability, hence things like spreasheet to csv and csv to double ArrayList.

## Rules
- If you change The value of the cell on first line called 'Note'. Be careful to change it inside the .java !

## Install
1. Copy the spreadsheet in your Google Drive (File > Make A Copy).
2. Fill the spreasheet with the marks.
3. In the Java Project:
Don't forget to provide a config.properties in a folder resources. It should be like this:
  host=theSMTPHost
  username=smtpUsername
  password=thePassword
4. InSendMails.java : Change the value of the google spreadhshet key.
