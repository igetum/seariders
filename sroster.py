import openpyxl
import csv
import re
import time

with open('extract-6.csv') as csv_file:

    csv_reader = csv.reader(csv_file, delimiter=",")

    teacherData = {}

    studentData = {}

    for data in csv_reader:
        if data and data[0].isdigit():
            studentid = data[0]
            student_last = data[1]
            student_first = data[2]
            term_start = data[4]
            term_end = data[5]
            teacher = data[6]
            period_start = data[7]
            grade = data[8]
            room = data[11]
        

            email_first = ''.join(e for e in student_first if e.isalnum())
            email_last = ''.join(e for e in student_last if e.isalnum())

            student_email = email_first + '.' + email_last + "@seariders.k12.hi.us"
            student_pass = student_first[-2:] + studentid[-6:]

            class_dict = {'teacher' : teacher, 'room' : room, 'period_start': period_start, 'term_start': term_start, 'term_end': term_end}

            if studentid not in studentData:
                classes = []
                classes.append(class_dict)
                studentData.setdefault(studentid, {'lastname' : student_last, 'firstname' : student_first, 'grade' : grade , 'email' : student_email.lower(), 'password' : student_pass, 'classes' : classes })
            else:
                classes = studentData[studentid]['classes']
                classes.append(class_dict)
                studentData[studentid]['classes'] = classes

            if teacher is '':
                teacher = 'activity'
           
            teacherData.setdefault(teacher, {})
            teacherData[teacher].setdefault(studentid, {'lastname' : student_last, 'firstname' : student_first, 'grade' : grade , 'email' : student_email.lower(), 'password' : student_pass, 'period_start' : period_start, 'room': room, 'term_start': term_start, 'term_end' : term_end})
            

csv_file.close()

wb = openpyxl.Workbook()

sheetname1 = "All Students"
sheetname2 = "Schedules"
wb.create_sheet(index=0, title=sheetname1)
wb.create_sheet(index=0, title=sheetname2)


########################### "ALL STUDENTS" ##################################
print("Proccessing All Students.... ")
sheet = wb[sheetname1]
sheet.print_area = 'A1:H27'
sheet['A1'] = 'Student Number'
sheet['B1'] = 'Last Name'
sheet['C1'] = 'First Name'
sheet['D1'] = 'Grade'
sheet['E1'] = 'Seariders Gmail'
sheet['F1'] = 'Password'

rowNum = 2

for studentid in studentData:
    col1 = sheet.cell(row=rowNum, column=1)
    col2 = sheet.cell(row=rowNum, column=2)
    col3 = sheet.cell(row=rowNum, column=3)
    col4 = sheet.cell(row=rowNum, column=4)
    col5 = sheet.cell(row=rowNum, column=5)
    col6 = sheet.cell(row=rowNum, column=6)
    
    firstname = studentData[studentid]['firstname']
    lastname = studentData[studentid]['lastname']
    grade = studentData[studentid]['grade']
    email = studentData[studentid]['email']
    password = studentData[studentid]['password']

    col1.value = studentid
    col2.value = lastname
    col3.value = firstname
    col4.value = grade
    col5.value = email
    col6.value = password

    rowNum += 1 

############################ "SCHEDULES" ##################################
print("Proccessing Schedule Sheet.... ")
sheet = wb[sheetname2]
sheet.print_area = 'A1:H27'
sheet['A1'] = 'Student Number'
sheet['B1'] = 'Last Name'
sheet['C1'] = 'First Name'
sheet['D1'] = 'Grade'
sheet['E1'] = 'Teacher'
sheet['F1'] = 'Room Name'
sheet['G1'] = 'Period Start'
sheet['H1'] = 'Seariders Gmail'
sheet['I1'] = 'Password'
sheet['J1'] = 'Term Start'
sheet['K1'] = 'Term End'
sheet['L1'] = 'sort(Term Start)'

rowNum = 2


for studentid in studentData:

    classlist = studentData[studentid]['classes']
    for studentclass in classlist:
        col1 = sheet.cell(row=rowNum, column=1)
        col2 = sheet.cell(row=rowNum, column=2)
        col3 = sheet.cell(row=rowNum, column=3)
        col4 = sheet.cell(row=rowNum, column=4)
        col5 = sheet.cell(row=rowNum, column=5)
        col6 = sheet.cell(row=rowNum, column=6)
        col7 = sheet.cell(row=rowNum, column=7)
        col8 = sheet.cell(row=rowNum, column=8)
        col9 = sheet.cell(row=rowNum, column=9)
        col10 = sheet.cell(row=rowNum, column=10)
        col11 = sheet.cell(row=rowNum, column=11)
        col12 = sheet.cell(row=rowNum, column=12)
    
        
        firstname = studentData[studentid]['firstname']
        lastname = studentData[studentid]['lastname']
        grade = studentData[studentid]['grade']
        email = studentData[studentid]['email']
        teacher = studentclass['teacher']
        room = studentclass['room']
        period_start = studentclass['period_start']
        password = studentData[studentid]['password']
        term_start = studentclass['term_start']
        term_end = studentclass['term_end']
        


        col1.value = studentid
        col2.value = lastname
        col3.value = firstname
        col4.value = grade
        col5.value = teacher
        col6.value = room
        col7.value = period_start
        col8.value = email
        col9.value = password
        col10.value = term_start
        col11.value = term_end
        col12.value = term_start[1]

        rowNum += 1 



############################  TEACHERS ###################################
print("Processing Teacher Sheets...")

for teacher in teacherData:

    wb.create_sheet(title=teacher)

    print("Processing Teacher Sheets [" + teacher + "]...")

    sheet = wb[teacher]

    sheet['A1'] = 'Student Number'
    sheet['B1'] = 'Last Name'
    sheet['C1'] = 'First Name'
    sheet['D1'] = 'Grade'
    sheet['E1'] = 'Teacher'
    sheet['F1'] = 'Room Name'
    sheet['G1'] = 'Period Start'
    sheet['H1'] = 'Seariders Gmail'
    sheet['I1'] = 'Password'
    sheet['J1'] = 'Term Start'
    sheet['K1'] = 'Term End'
    sheet['L1'] = 'sort(Term Start)'


    rowNum = 2

    for studentid in teacherData[teacher]:
        student = teacherData[teacher][studentid]

        col1 = sheet.cell(row=rowNum, column=1)
        col2 = sheet.cell(row=rowNum, column=2)
        col3 = sheet.cell(row=rowNum, column=3)
        col4 = sheet.cell(row=rowNum, column=4)
        col5 = sheet.cell(row=rowNum, column=5)
        col6 = sheet.cell(row=rowNum, column=6)
        col7 = sheet.cell(row=rowNum, column=7)
        col8 = sheet.cell(row=rowNum, column=8)
        col9 = sheet.cell(row=rowNum, column=9)
        col10 = sheet.cell(row=rowNum, column=10)
        col11 = sheet.cell(row=rowNum, column=11)
        col12 = sheet.cell(row=rowNum, column=12)
        

        col1.value = studentid
        col2.value = student['lastname']
        col3.value = student['firstname']
        col4.value = student['grade']
        col5.value = student['period_start']
        col6.value = student['room']
        col7.value = student['period_start']
        col8.value = student['email']
        col9.value = student['password']
        col10.value = student['term_start']
        col11.value = student['term_end']
        col12.value = student['term_start'][1]


        rowNum += 1


wb.save("output.xlsx")
 
