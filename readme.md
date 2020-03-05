# Sending e-mails app
#### Dedicated app sending e-mails with formatted body and special attachments.

<img src="main-screen.png" width="700">

## Table of contents:
#### * General info
#### * Technologies 
#### * Setup
#### * Status

## **General info**:
The main goal of this app is reducing time spent on sending specific e-mails during my working in small family's company. Thanks to it I made up a 15 minutes every day and decreased probability of making mistakes.

The app looks like a notepad with a crucial tool kit and set of necesseries buttons.
The main functions are:
* uploading archival .txt file (required for updating data)
* overwriting this file
* exporting newly part of uploaded .txt file into .csv 
* sending e-mail with: 
     - weekday-related title
     - formatted body using different colors depending on content of exported .csv file
     - attached particular attachments. After sending them the app moves sent attachments to a local archive directory.
     
Additionally, the app shows warning boxes in cases like exiting without saving or overwriting file.
     
## **Technologies**:
- Python 3.7.4
- HTML 

### Libraries and packages which have been used:
 - os
 - encodings
 - smtplib
 - calendar
 - fnmatch
 - shutil
 - email
 - datetime
 - **csv**
 - **PIL**
 - **pandas**
 - **tkinter** (as only GUI framework)
 
 ## **Setup**:
 macOS 10.14.6
 
 As I mentioned in "General info" section, this app fulfills very specific requirements. However, in case of desire to solve  similar problems please follow below steps:
 
 Once you copy this repo on your local computer please install requirements.txt by entering in terminal: ```pip install -r requirements.txt```.
 
 The app uses specific paths, file names or email addresses. Please add yours in the file: ```private.xlsx```.
 
 <img src="privatexlslx-screen.png" width="900">
 
 Then go on to the following part:
 1. After running the app ```GIT_emailFV.py``` the notepad window will appear.
 2. Open archival text file by clicking on first on the left button. *If you didn't specify path to file in private.xlsx the default archival text file is ```data.txt```.*
 Now we can see the most current archival version of this file. 
 3. Go down to the end of the file and write new data into notepad bearing in mind proper syntax (cf. begining of file).
 4. Overwrite the notepad by clicking on second on the left button or choosing appropriate icon in "File" menubar.
 5. Export file to CSV by clicking on third on the left button or choosing appropriate icon in "File" menubar. Please select path where ```export.csv``` is located. 
  <img src="export-screen.png" width="700">
  
 6. The only thing to do is clicking "Save".
 
 7. Send e-mail by clicking on the last button. Your email's receiver will see email which looks like:
  <img src="email-screen.png" width="700">

## **Status**:

The next steps will be:
1. adding a window converting foreign exchange rates
2. transfer excel database into sql one and keep using related database for further work
