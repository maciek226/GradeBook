# GradeBook
## Feature Highlights 
The app was intended for simplifying the grade input process for individuals unlucky enough to have to use blackboard services, however, it may find use in other applications as well.
- Searching multiple spreadsheets 
- Partial matches of both names and student numbers
- Entering multiple grades for the same student at the same time 

## Overview 
If your school is using blackboard and you are tasked with entering grades the process will take forever, especially if the course has many sections, this script may reduce the extent of your misery. This app opens the spreadsheets downloadable from blackboard and allows you to quickly search for a student by portion of first name, last name, or student number, and enter the grade. The script supports searching multiple spreadsheets and entering multiple grades for a given student, which is useful for creating mark breakdowns for accreditation. 
**Use the app at your own risk.**
The complied version of the app was tested on Windows 10. The pyton script should work on other platforms. 

# Quick Start  
1. Download the application executable (GradeBook.exe) from github 
(check repository for the newest version: https://github.com/maciek226/GradeBook)
2. Download the spreadsheets with students names and desired grades from blackboad
3. Open GradesConf.txt and add add the path to the spreadsheets you would like to search 
4. For each sheet start input "in" followed by a tab and the path to the file 
5. Open GradeBook.exe
6. Follow instructions in the command line window

## Commands
The script operates using written commands that differ slightly depending on the task. Note, the commands in the script are case sensitive

- `EXT` - Exit and save\
  Available only in the search for student filed 
- `CNC` - Cancel\
  Available whenever there is a choice such as adding a student grade, selecting student from the list, or adding spreadsheet column
- `ADD`- Add column\
  Available only in the column selection screen 

## Adding a Column to a Spreadsheet 
The app can modify the size of the spreadsheet using `ADD` command. This feature is intended to aid in creating mark breakdowns (needed for accreditation). These new fields, however, will not be read by blackboard upon upload of the updated sheet. Therefore, add all the grade fields before downloading the spreadsheet to your computer. Before uploading the updated spreadsheet make a copy of the file for your records and delete the added columns.

## Limitations 
- The script can catch serious errors; however, it does not have any logic checking ability. 
- Make sure that when selecting columns in multiple sheets the columns are selected in the same order for all sheets. 
- It is not possible to change any configurations (set up of columns) after the selection.
- The grades must be numbers (natural or positive real) 

## Configuration File
- Configuration file is a tab separated file (tab separated the property name from value)
- Add spreadsheets by writing int followed by tab (`\t`) and the path to the file\
  `in \t  C:\User\...`
- The modifications to the file can be either overwritten by setting\
  `overwrite \t 1`
- To write data to a new file set\
  `overwrite \t 0` and add an output file path\
  `out \t C:User\`
- If the configuration file is missing the script will ask for the paths to the desired files. Once complete, a new config file will be made in the root directory

## Troubleshooting 
- Make sure that the spreadsheets you are editing are not open in other applications \
  The script can detect if other app prevents it from closing and it should give a warning

## Other Information
- The script is built in Python 3
