# excel file reader
import pandas as pd
# to find numbers within strings
import re
# excel file writer
import xlsxwriter
# to retrieve file extension information
import os
# to convert ascii to float
import locale 

# tkinter GUI
from tkinter import *
root = Tk()

from tkinter import ttk
from tkinter import filedialog as fd

root.title('Roster Viewer')
root.geometry("1570x1000")
# icon
try:
    img = PhotoImage(file='./_internal/Grad.png')
    root.iconphoto(False, img)
except:
    print("There is no icon")

# width and height of Ledger_Viewer
RVwidth = 1570
RVheight = 1000
header = 75

# add some style
style = ttk.Style()
style.theme_use('clam')
style.map('Treeview')

my_tree = ttk.Treeview(root)

# Define our columns
my_tree['columns'] = ("Term", "Dept", "Course", "Title", "Instructor", "CRs",
                      "Bldg", "Room", "Days", "Start Time", "End Time", "Enrlmnt", "Student ID",
                      "Student", "Registered", "Start Term", "Grade", "Major Type", "Major", "Cl",
                      "Email", "Registration Status")
# Format our columns
CW = int(.00318*RVwidth)
my_tree.column("#0", width=40, minwidth=8*CW)
my_tree.column("Term", anchor=W, width=10*CW)
my_tree.column("Dept", anchor=W, width=CW)
my_tree.column("Course", anchor=W, width= 15*CW)
my_tree.column("Title", anchor=W, width= 40*CW)
my_tree.column("Instructor", anchor=W, width= 30*CW)
my_tree.column("CRs", anchor=W, width= 8*CW)
my_tree.column("Bldg", anchor=W, width= 8*CW)
my_tree.column("Room", anchor=W, width= 8*CW)
my_tree.column("Days", anchor=W, width= 8*CW)
my_tree.column("Start Time", anchor=W, width= 12*CW)
my_tree.column("End Time", anchor=W, width= 12*CW)
my_tree.column("Enrlmnt", anchor=W, width= 12*CW)
my_tree.column("Student ID", anchor=W, width= 16*CW)
my_tree.column("Student", anchor=W, width= 30*CW)
my_tree.column("Registered", anchor=W, width= CW)
my_tree.column("Start Term", anchor=W, width= CW)
my_tree.column("Grade", anchor=W, width= 8*CW)
my_tree.column("Major Type", anchor=W, width= CW)
my_tree.column("Major", anchor=W, width= 20*CW)
my_tree.column("Cl", anchor=W, width= 8*CW)
my_tree.column("Email", anchor=W, width= 48*CW)
my_tree.column("Registration Status", anchor=W, width= CW)

# even and odd row coloring
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")
# red coloring for negative transactions
my_tree.tag_configure('positive', foreground="black")
my_tree.tag_configure('negative', foreground="red2")

# define headings
my_tree.heading("#0", text="Roster", anchor=W)
my_tree.heading("Term", text="Term", anchor=W)
my_tree.heading("Dept", text="Dept", anchor=W)
my_tree.heading("Course", text="Course", anchor=W)
my_tree.heading("Title", text="Title", anchor=W)
my_tree.heading("Instructor", text="Instructor", anchor=W)
my_tree.heading("CRs", text="CRs", anchor=W)
my_tree.heading("Bldg", text="Bldg", anchor=W)
my_tree.heading("Room", text="Room", anchor=W)
my_tree.heading("Days", text="Days", anchor=W)
my_tree.heading("Start Time", text="Start Time", anchor=W)
my_tree.heading("End Time", text="End Time", anchor=W)
my_tree.heading("Enrlmnt", text="Enrlmnt", anchor=W)
my_tree.heading("Student ID", text="Student ID", anchor=W)
my_tree.heading("Student", text="Student", anchor=W)
my_tree.heading("Registered", text="Registered", anchor=W)
my_tree.heading("Start Term", text="Start Term", anchor=W)
my_tree.heading("Grade", text="Grade", anchor=W)
my_tree.heading("Major Type", text="Major Type", anchor=W)
my_tree.heading("Major", text="Major", anchor=W)
my_tree.heading("Cl", text="Cl", anchor=W)
my_tree.heading("Email", text="Email", anchor=W)
my_tree.heading("Registration Status", text="Registration Status", anchor=W)

# add a scrollbar for viewing
vsb = ttk.Scrollbar(root, orient="vertical", command=my_tree.yview)
vsb.place(x=RVwidth-20, y=header, height=RVheight-header-30)  
my_tree.configure(yscrollcommand=vsb.set)

# list of students 
complete_roster = []

def resizer(event):
    # only resize on main window event
    if(event.widget.master == None):
        # width and height of Ledger_Viewer
        RVwidth = event.width
        RVheight = event.height
        header = 75

        # Format our columns
        CW = int(.00318*RVwidth)
        my_tree.column("#0", width=40, minwidth=8*CW)
        my_tree.column("Term", anchor=W, width=10*CW)
        my_tree.column("Dept", anchor=W, width=CW)
        my_tree.column("Course", anchor=W, width= 15*CW)
        my_tree.column("Title", anchor=W, width= 40*CW)
        my_tree.column("Instructor", anchor=W, width= 30*CW)
        my_tree.column("CRs", anchor=W, width= 8*CW)
        my_tree.column("Bldg", anchor=W, width= 8*CW)
        my_tree.column("Room", anchor=W, width= 8*CW)
        my_tree.column("Days", anchor=W, width= 8*CW)
        my_tree.column("Start Time", anchor=W, width= 12*CW)
        my_tree.column("End Time", anchor=W, width= 12*CW)
        my_tree.column("Enrlmnt", anchor=W, width= 12*CW)
        my_tree.column("Student ID", anchor=W, width= 16*CW)
        my_tree.column("Student", anchor=W, width= 30*CW)
        my_tree.column("Registered", anchor=W, width= CW)
        my_tree.column("Start Term", anchor=W, width= CW)
        my_tree.column("Grade", anchor=W, width= 8*CW)
        my_tree.column("Major Type", anchor=W, width= CW)
        my_tree.column("Major", anchor=W, width= 20*CW)
        my_tree.column("Cl", anchor=W, width= 8*CW)
        my_tree.column("Email", anchor=W, width= 48*CW)
        my_tree.column("Registration Status", anchor=W, width= CW)
        my_tree.place(x = 10, y = header, width = RVwidth-30, height = RVheight-header-10)

        # change the location and size of the scrollbar
        vsb.place(x=RVwidth-20, y=header, height=RVheight-header-10)  

        CR_label.place(x=RVwidth-250, y=10)
        CRbox.place(x=RVwidth-185, y=10)
        check.place(x=RVwidth-115, y=10)

        ID_label.place(x=RVwidth-250, y = 43)
        IDbox.place(x=RVwidth-185, width=100, y=43)
        IDbutton.place(x=RVwidth-75, y=38, width=60, height=30)

def print_roster(event):
    k = 0
    child = 0
    first_row_in_class = 0
    even_odd = 0
    prev_class = ''
 
    students = []
    instructors = []
    majors = []

    #clear display
    my_tree.delete(*my_tree.get_children())

    #Fill student dropdown
    if(drop_dept.get() != "" and IDbox.get() == ''):
        drop_student.delete(0, END)
        for row in complete_roster:
            if drop_dept.get() == row[1]:
                original = True
                for i in students:
                    if(i == row[13]):
                        original = False
                        break;
                if(original == True):
                    students.append(row[13])
        students.sort()
        students.insert(0, "")
        drop_student['values'] = students
    else:
        drop_student.delete(0, END)
        students.insert(0, "")
        drop_student['values'] = students
        
    #Fill instructor dropdown
    if(drop_dept.get() != "" and IDbox.get() == ''):
        drop_instruct.delete(0, END)
        for row in complete_roster:
            if drop_dept.get() == row[1]:
                original = True
                for i in instructors:
                    if(i == row[4]):
                        original = False
                        break;
                if(original == True):
                    instructors.append(row[4])
        instructors.sort()
        instructors.insert(0, "")
        drop_instruct['values'] = instructors
    else:
        drop_instruct.delete(0, END)
        instructors.insert(0, "")
        drop_instruct['values'] = instructors

    #Fill major dropdown
    if(drop_dept.get() != "" and IDbox.get() == ''):
        drop_major.delete(0, END)
        for row in complete_roster:
            if drop_dept.get() == row[1]:
                original = True
                for i in majors:
                    if(i == row[18]):
                        original = False
                        break;
                if(original == True):
                    majors.append(row[18])
        majors.sort()
        majors.insert(0, "")
        drop_major['values'] = majors
    else:
        drop_major.delete(0, END)
        majors.insert(0, "")
        drop_major['values'] = majors
        
    if(IDbox.get() != ''):
        drop_dept.current(0)

    # remove duplicate rows
    modified_roster = []
    r = 0
    while r < len(complete_roster):
        # if we have the same student and class, it's a duplicate record
        if(r < len(complete_roster)-1):
            if(complete_roster[r][13] == complete_roster[r+1][13] and complete_roster[r][2] == complete_roster[r+1][2]):
                # column 8 is the meeting days, final exam days are duplicate records
                countrow = sum([1 for c in complete_roster[r][8] if c.isalpha()])
                countnextrow = sum([1 for c in complete_roster[r+1][8] if c.isalpha()])
                if(countrow > countnextrow):
                    modified_roster.append(complete_roster[r])
                else:
                    modified_roster.append(complete_roster[r+1])
                r += 2
            else:
                modified_roster.append(complete_roster[r])
                r += 1
        else:
            modified_roster.append(complete_roster[r])
            r += 1
    # some courses are in triplicate
    modified_roster2 = []
    r = 0
    while r < len(modified_roster):
        # if we have the same student and class, it's a duplicate record
        if(r < len(modified_roster)-1):
            if(modified_roster[r][13] == modified_roster[r+1][13] and modified_roster[r][2] == modified_roster[r+1][2]):
                # column 8 is the meeting days, final exam days are duplicate records
                countrow = sum([1 for c in modified_roster[r][8] if c.isalpha()])
                countnextrow = sum([1 for c in modified_roster[r+1][8] if c.isalpha()])
                if(countrow > countnextrow):
                    modified_roster2.append(modified_roster[r])
                else:
                    modified_roster2.append(modified_roster[r+1])
                r += 2
            else:
                modified_roster2.append(modified_roster[r])
                r += 1
        else:
            modified_roster2.append(modified_roster[r])
            r += 1
    CRs = 0
    for row in modified_roster2:
        print_r = False
        if(drop_dept.get() == '' and drop_student.get() == '' and 
           drop_instruct.get() == '' and drop_major.get() == '' and IDbox.get() == ''):
            print_r = True
        elif(drop_dept.get() == row[1] and drop_student.get() == '' and drop_instruct.get() == '' and drop_major.get() == ''):
            print_r = True
        elif(drop_instruct.get() == row[4]):
            print_r = True
        elif(drop_student.get() == row[13]):
            print_r = True
        elif(row[18] in drop_major.get() and drop_major.get != ''):
            print_r = True
        elif(row[12] in IDbox.get() and IDbox.get() != ''):
            print_r = True
 
        if(print_r == True):
            # get the course number
            num = re.findall(r'\d+', row[2])
            if((locale.atof(row[11]) < 10) and (locale.atof(num[0]) < 500)):
                color = 'negative'
            elif(locale.atof(row[11]) < 6):
                color = 'negative'
            else:
                color = 'positive'

            if((thesis.get() and (locale.atof(num[0]) == 595 or
                                   locale.atof(num[0]) == 596 or
                                   locale.atof(num[0]) == 599 or 
                                   locale.atof(num[0]) == 695 or 
                                   locale.atof(num[0]) == 699)) or 
                                   thesis.get() == 0):
                # if students are in the same class, dont' change the color
                if(row[2] == prev_class):
                    even_odd -= 1
                else:
                    header = row[:12]
                    if even_odd % 2 == 0:
                        my_tree.insert(parent='', index='end', iid=k, text="", values=header, tags=('evenrow',color))
                    else:
                        my_tree.insert(parent='', index='end', iid=k, text="", values=header, tags=('oddrow',color))
                    first_row_in_class = k
                    k += 1
                    child = 0
                
                student_record = ['']*22
                student_record[12:] = row[12:]
                if even_odd % 2 == 0:
                    my_tree.insert(parent='', index='end', iid=k, text="", values=student_record, tags=('evenrow',color))
                else:
                    my_tree.insert(parent='', index='end', iid=k, text="", values=student_record, tags=('oddrow',color))

                my_tree.move(k, first_row_in_class, child)
                even_odd += 1
                child += 1

                k+=1
                prev_class = row[2]
                CRs += locale.atof(row[5])
    # print the total number of credit hours for the semester
    CRbox.configure(state='normal')
    CRbox.delete(0,END)
    strval = "{:,.1f}".format(CRs)
    CRbox.insert(0,strval)
    CRbox.configure(state='disabled')

def update_dept(event):
    drop_student.current(0)
    drop_instruct.current(0)
    drop_major.current(0)
    IDbox.delete(0, END)
    print_roster(0)
def clear_dept(event):
    drop_dept.current(0)
    print_roster(0)
def update_student(event):
    drop_instruct.current(0)
    drop_major.current(0)
    IDbox.delete(0, END)
    print_roster(0)
def clear_student(event):
    drop_student.current(0)
    print_roster(0)
def update_instruct(event):
    drop_student.current(0)
    drop_major.current(0)
    IDbox.delete(0, END)
    print_roster(0)
def clear_instruct(event):
    drop_instruct.current(0)
    print_roster(0)
def update_major(event):
    drop_student.current(0)
    drop_instruct.current(0)
    IDbox.delete(0, END)
    print_roster(0)
def clear_major(event):
    drop_major.current(0)
    print_roster(0)
def updateID(event):
    drop_student.current(0)
    drop_instruct.current(0)
    drop_major.current(0)
    drop_dept.current(0)
    print_roster(0)
def clear_IDbox(event):
    IDbox.delete(0, END)
    print_roster(0)

# define dept combo box
drop_dept = ttk.Combobox(root,values = "", state="readonly", width=25)
dept_label = Label(root, text="Dept.")
dept_label.place(x=210, y=10)
drop_dept.place(x=250, y=10)
drop_dept.bind("<<ComboboxSelected>>", update_dept)
drop_dept.bind("<Button-3>", clear_dept)
drop_dept.bind("<Delete>", clear_dept)

# define instructor combo box
drop_instruct = ttk.Combobox(root,values = "", state="readonly", width=25)
instruct_label = Label(root, text="Instructor")
instruct_label.place(x=450, y=10)
drop_instruct.place(x=510, y=10)
drop_instruct.bind("<<ComboboxSelected>>", update_instruct)
drop_instruct.bind("<Button-3>", clear_instruct)
drop_instruct.bind("<Delete>", clear_instruct)

# define student combo box
drop_student = ttk.Combobox(root,values = "", state="readonly", width=25)
student_label = Label(root, text="Student")
student_label.place(x=450, y=43)
drop_student.place(x=510, y=43)
drop_student.bind("<<ComboboxSelected>>", update_student)
drop_student.bind("<Button-3>", clear_student)
drop_student.bind("<Delete>", clear_student)

# define major combo box
drop_major = ttk.Combobox(root,values = "", state="readonly", width=25)
major_label = Label(root, text="Major")
major_label.place(x=210, y=43)
drop_major.place(x=250, y=43)
drop_major.bind("<<ComboboxSelected>>", update_major)
drop_major.bind("<Button-3>", clear_major)
drop_major.bind("<Delete>", clear_major)

# define student ID combo box
xplace = 10
IDbox = ttk.Entry(root)
ID_label = Label(root, text="Student ID")
ID_label.place(x=RVwidth-250, y = 43)
IDbox.place(x=RVwidth-185, width=100, y=43)
IDbox.bind("<Button-3>", clear_IDbox)
IDbox.bind("<Delete>", clear_IDbox)

# define a search button
IDbutton = ttk.Button(root, text="Search", command= lambda:updateID(0))
IDbutton.place(x=RVwidth-75, y=38, width=60, height=30)

# define CR tally
CRbox = ttk.Entry(root, width=10)
CRbox.config(foreground="black")

CR_label = Label(root, text="Total CRs")
CR_label.place(x=RVwidth-250, y=10)
CRbox.place(x=RVwidth-185, y=10)

# define thesis checkbox
thesis = IntVar()
check = Checkbutton(root, text="Thesis Only",bd=0, variable = thesis, onvalue=1,
                       offvalue=0, command = lambda: print_roster(0))
check.place(x=RVwidth-115, y=10)


# double-click on items to copy them to clipboard
def select(event):
    coltext = my_tree.identify_column(event.x)
    col = re.findall(r'\d+', coltext)
    column = int(col[0])-1
    selected = my_tree.selection()
    text = str('')
    for row in selected:
        values = my_tree.item(row, 'values')
        text = text + str(values[column]) + '\n'
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()
my_tree.bind('<Double-1>', select)

# selection of rows
def sum_rows(event):
    if(len(my_tree.selection()) > 0):
        CRs = 0
        for item in my_tree.selection():
            record = my_tree.item(item, 'values')
            if(record[5] != ''):
                CRs += locale.atof(record[5])*locale.atof(record[11])
        # print the total number of credit hours for the semester
        CRbox.configure(state='normal')
        CRbox.delete(0,END)
        CRbox.insert(0,str(CRs))
        CRbox.configure(state='disabled')

my_tree.bind("<<TreeviewSelect>>", sum_rows)

#bind GUI resizing
root.bind('<Configure>', resizer)

root.bind('<Button-3>', print_roster)
root.bind('<Return>', print_roster)

# select a file from the list
def select_file():
    filetypes = (
        ('excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    getfile = fd.askopenfilename(
        title='Open File',
        initialdir='.',
        filetypes=filetypes)

    if(getfile == ''):
        return

    # get transactions from csv file
    if(".xlsx" in getfile):
        with pd.ExcelFile(getfile) as excl:
            sheets = excl.sheet_names
            df = excl.parse(sheets[0])
            matrix = df.to_numpy()
            complete_roster.clear()

            for listrow in matrix:
                row = list(map(str,listrow))
                try:
                    if(row[0].isnumeric()):
                        complete_roster.append(row)
                except:
                    continue
        
            # fill in the department dropdown
            depts = []
            for listrow in matrix:
                row = list(map(str,listrow))
                try:
                    if(row[0].isnumeric()):
                        original = True
                        for i in depts:
                            if(i == row[1]):
                                original = False
                                break
                        if(original == True):
                            depts.append(row[1])
                except:
                    continue
            depts.sort()
            depts.insert(0,"")
            drop_dept.set('')
            drop_dept['values'] = depts

            # clear the dropdowns
            drop_instruct.set('')
            drop_student.set('')
            drop_major.set('')

            root.title(getfile)
 
            print_roster(0)

def save_file():
    title = ['Term', 'Dept', 'Course', 'Title', 'Instructor', 'CRs', 'Bldg', 'Room', 'Days', 
             'Start Time', 'End Time', 'Enrlmnt', 'Student ID', 'Student', 'Registered', 
             'Start Term', 'Grade', 'Major Type', 'Major', 'Cl', 'Email', 'Registration Status']

    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    csvfile = fd.asksaveasfilename(
        title='Save file',
        initialdir='.',
        filetypes=filetypes)
    fn = os.path.splitext(csvfile)
    if(fn[0] != ''):
        # add file extension if not specified
        fext = fn[1]
        if(fext == ''):
            fext = '.xlsx'
            csvfile += fext

        if(fext == ".xlsx"):
            workbook = xlsxwriter.Workbook(csvfile)
            sheet = workbook.add_worksheet()
            for idx_col, col in enumerate(title):
                sheet.write(0, idx_col, col)
            idx_row = 1
            for row in my_tree.selection():
                values = my_tree.item(row, 'values')
                for idx_col, col in enumerate(values):
                    sheet.write(idx_row, idx_col, col)
                idx_row += 1
                prev_row = values
            workbook.close()

# open button
open_button = ttk.Button(
    root,
    text='Open File',
    command=select_file
)
# put it in the upper left corner
open_button.place(x=10, y=5)

# save button
save_button = ttk.Button(
    root,
    text='Save Selected',
    command=save_file
)
# put it in the upper left corner
save_button.place(x=100, y=5)

root.mainloop()
