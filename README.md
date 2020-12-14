# Zoom-Attendance-Script
## General Information
* Uses Zoom generated attendance excel files to create a master attendance sheet
* Iterates over the names of the attendance files and searches for the name in the master file. If a name is found their attendance count will increase. If a name is not found the name will be added and their attendance count will be 1. 

## Libraries Used
* To run please install the following libraries
* pip install xlrd
* pip install xlsxwriter

## Run
* Must input the file path of the excel file being inputted and file path of the master attendence file
