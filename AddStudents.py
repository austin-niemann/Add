import pandas as pd
import xlsxwriter as xlsxwriter
import subprocess, sys

filePath = input("Enter File Path")
user = input("enter first.last name")
# import student name csv file
new = filePath[1:-1]

df = pd.read_csv(r"%s" % new, header=None, encoding='UTF-8')

# drop blank lines

df.dropna(axis=0, how="any", thresh=None, subset=None, inplace=True)

# find the number of lines without blanks

size = df.shape
print(new)


# number of students to load (determines how many times the while loops run)

students = max(size)
# write new excel document
docName = "Load.xlsx"
newPath = r'C:\Users\%s\Desktop\Load.xlsx' % user
workbook = xlsxwriter.Workbook(newPath)
worksheet = workbook.add_worksheet()

# column headers
worksheet.write('A1', "FIRSTNAME")
worksheet.write('B1', "LASTNAME")
worksheet.write('C1', "USERNAME")
worksheet.write('D1', "PASSWORD")
worksheet.write('E1', "OU")
worksheet.write('F1', "DESCRIPTION")
worksheet.write('G1', "Principal Name")


# constant variables
OU = "OU=WLC,OU=1st Battalion,DC=rti,DC=loc"
Description = "BLC Student"
Password = "password"

# while loop to input students first names


condition_First = 0
intRow_First = 0
intColumn_First = 0
intCell_First = 2
firstName = (df.iloc[intRow_First, intColumn_First])
print(firstName)


while condition_First < students:
    firstName = (df.iloc[intRow_First, intColumn_First])
    worksheet.write('A%d' % intCell_First, firstName)

    intCell_First += 1
    intRow_First += 1
    intColumn_First += 0
    condition_First += 1

# while loop to input students last names, OU, Password, and Description

condition_last = 0
intRow_last = 0
intColumn_last = 1
intCell_last = 2
intCell_OU = 2
intCell_Desc = 2
intCell_Pass = 2


while condition_last < students:
    lastName = (df.iloc[intRow_last, intColumn_last])
    worksheet.write('B%d' % intCell_last, lastName)
    worksheet.write('E%d' % intCell_OU, OU)
    worksheet.write('F%d' % intCell_Desc, Description)
    worksheet.write('D%d' % intCell_Pass, Password)

    condition_last += 1
    intRow_last += 1
    intColumn_last += 0
    intCell_last += 1
    intCell_OU += 1
    intCell_Desc += 1
    intCell_Pass += 1


# while loop to input students username

condition_user = 0
intRow_user = 0
intColumn_user = 2
intCell_user = 2
intCell_PN = 2

while condition_user < students:
    username = (df.iloc[intRow_user, intColumn_user])
    worksheet.write('C%d' % intCell_user, username)
    worksheet.write('G%d' % intCell_PN, "%s@rti.loc" % username)
    print("added %s" % username)

    condition_user += 1
    intRow_user += 1
    intColumn_user += 0
    intCell_user += 1
    intCell_PN += 1

# while loop to input password

#condition_pass = 0
#intRow_pass = 0
#intColumn_pass = 1
#intCell_pass = 2

#while condition_pass < students:
    #password = (df.iloc[intRow_pass, intColumn_pass])
    #worksheet.write('D%d' % intCell_pass, password)

    #condition_pass += 1
    #intRow_pass += 1
    #intColumn_pass += 0
    #intCell_pass += 1
# end of while loops and finishes writing excel document
workbook.close()
print("File saved to desktop as %s" % docName)

#p = subprocess.Popen(['powershell.exe', r"C:\Users\austin.niemann\Desktop\powershelladmin.ps1"], stdout=sys.stdout)
#p.communicate()

print("Added %d new students to Active Directory!" % students)







