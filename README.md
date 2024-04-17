# JMU-Fitness-Club
This tool helps streamline the post-meeting attendance and strike paperwork. This README is a work in progess.

## Setting up pygsheets
If needed, follow [this tutorial](https://www.geeksforgeeks.org/how-to-automate-google-sheets-with-python/) to set up the Google Sheets API to get a service account

## Setting up Google Sheets
This system ONLY functions if Google Sheets is formatted in the manner shown below (If you wish to have a different format, you must change the hardcoded column and row numbers in `strikes_rev2.py`. The names of the spreadsheet itself and its pages sare customizable, but the user should be sure to change the names of these within the code before use. The "All Present & Excused Absent," "All Excused Late," and "All Unexcused Absent" columns will be populated by the service account as the user is running the Python program. In the event that you need to edit the contents of these  columns, do not do so until the program has finished running. When inputting what people are excused absent/late, you should follow the same format as shown in the below examples. Capitalization does not matter, but make sure the the words are EXACTLY spelled as shown, or else the system may incorrectly assign the corresponding person an absense or strike.

### Spreadsheet Layout
<img src = "Spreadsheet Examples/Strikes Spreadsheet Example.png" />
<img src = "Spreadsheet Examples/Attendance Example.png" />

