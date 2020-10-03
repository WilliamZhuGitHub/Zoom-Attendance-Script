""""
Libraries:
=====================================================================================
Type the following commands in the comments in cmd
pip install xlrd
pip install xlsxwriter
=====================================================================================
"""
import xlrd #pip install xlrd
import xlsxwriter #pip install xlsxwriter
""""
INSTRUCTIONS for FILE PATH: 
=====================================================================================
Drag excel spread sheet into the python folder(Make sure the file type is xlsx not csv)
Paste the file path to variable "todayAttendance"
Need r in front of path
Format should be in the form of (r'"insert file path here"')
file paths are the only things that need to be modified per use.
=====================================================================================
"""

#file path to today's attendance(should be changed every time this is ran)
todayAttendance = (r'C:\Users\zhuw2\Documents\GitHub\Zoom-Attendance-Script\9_21.xlsx')

#file path to total attendance list(should not be changed after first time set)
totalAttendance = (r'C:\Users\zhuw2\Documents\GitHub\Zoom-Attendance-Script\Total Attendance.xlsx')


#open today's attendance data
today_workbook = xlrd.open_workbook(todayAttendance)
today_worksheet = today_workbook.sheet_by_index(0)

#open total attendance data
total_workbook = xlrd.open_workbook(totalAttendance)
total_worksheet = total_workbook.sheet_by_index(0)

all_rows = []

#iterates through the first column of today's attendance spreadsheet and adds the cell contents to an array
#return an array of all names from today's attendance
def arrayOfAttendance():
    all_names_today = []
    for row in range(today_worksheet.nrows):
        all_names_today.append(today_worksheet.cell_value(row,0))
    return all_names_today

#iterates through the first column of the total attendance spreadsheet and adds the cell contents to an array
#return an array of all names from the total attendance sheet
def arrayOfAllNames():
    all_names = []
    for row in range (total_worksheet.nrows):
        all_names.append(total_worksheet.cell_value(row, 0))
    return all_names
 

#check if new members has joined 
#iterates through the list of today's attendance
#checks if the name exists on the total attendance sheet
#if it doesn't it adds the name to the attendance sheet with attendance set to 1
#will print to console when a new name has been added
def ensureNames():
    for i in (arrayOfAttendance()):
        add_row = []
        if(i not in arrayOfAllNames()):
            add_row.append(i)
            add_row.append(1)
            all_rows.append(add_row)
            print("added a new name: " + i)
    return

#copy data from total attendance sheet into a 2-D array 
for row in range(total_worksheet.nrows):
    curr_row = []
    for col in range(total_worksheet.ncols):
        #skips the header cell and the name column
        if(col == 1 and row > 0): 
            curr_row.append(int(total_worksheet.cell_value(row,col)))
        else:
            curr_row.append(total_worksheet.cell_value(row,col))
    all_rows.append(curr_row)
 
#add 1 to attendance in the 2-D array if the name is present in attendance and total 
for row in range(total_worksheet.nrows):
    if((total_worksheet.cell_value(row,0)) in arrayOfAttendance() and row > 0):
            all_rows[row][1]  += 1

#calls the ensureName method after attendance has been checked
ensureNames()

#write the data from the all_rows 2-D array back to excel
writeTototal = xlsxwriter.Workbook(totalAttendance)
writeTototalSheet = writeTototal.add_worksheet()

for row in range (len(all_rows)):
    for col in range(len(all_rows[0])):
        writeTototalSheet.write(row,col,all_rows[row][col])

writeTototal.close()

#output to console when completed
print("Success. File: " + todayAttendance + " was added")

 

