# The script was created by Oleg Timerman in september 2021
# The purpose of this script is to create the list of files in directory and sub directoryes.
# The file names will be writen to new excel file named index.xlsx.
# First column name of file and the second column is a hyperlink. (Full path to file)

#!/usr/bin/python3
import os
import xlsxwriter

def writefilepath():
    path= os.getcwd()
#creating workbook and worksheet
    workbook = xlsxwriter.Workbook("index.xlsx")
    sheet = workbook.add_worksheet()
    y = 1
    sheet.write(0, 0, "file name")
    sheet.write(0, 1, "file path")
# creating the list of pathfiles to "filepath"
    for root, dirs, files in os.walk(path):
        for filenames in files:
# writing the filepath to workbook
            sheet.write_url(y, 1, os.path.join(root, filenames))
            sheet.write(y, 0, filenames)
            y = y + 1
    workbook.close()
def main():
    writefilepath()
    print("Thank you it's DONE !!!")
    input("Please input enter to exit ")

if __name__ == "__main__":
    main()