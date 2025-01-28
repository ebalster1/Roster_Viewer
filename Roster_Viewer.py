# excel file reader
import pandas as pd
import re

# tkinter GUI
from tkinter import *
root = Tk()

from tkinter import ttk
from tkinter import filedialog as fd

root.title('Roster Viewer')
root.geometry("1570x1000")
# width and height of Ledger_Viewer
RVwidth = 1570
RVheight = 1000
header = 50
#root.resizable(0,0)

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
my_tree.column("Student ID", anchor=W, width= 4*CW)
my_tree.column("Student", anchor=W, width= 30*CW)
my_tree.column("Registered", anchor=W, width= CW)
my_tree.column("Start Term", anchor=W, width= CW)
my_tree.column("Grade", anchor=W, width= 8*CW)
my_tree.column("Major Type", anchor=W, width= CW)
my_tree.column("Major", anchor=W, width= 20*CW)
my_tree.column("Cl", anchor=W, width= 8*CW)
my_tree.column("Email", anchor=W, width= 60*CW)
my_tree.column("Registration Status", anchor=W, width= CW)

# even and odd row coloring
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")
# red coloring for negative transactions
my_tree.tag_configure('positive', foreground="black")
my_tree.tag_configure('negative', foreground="red2")

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
vsb.place(x=RVwidth-13, y=header, height=RVheight-header-30)  
my_tree.configure(yscrollcommand=vsb.set)

complete_roster = []

def resizer(event):
    # only resize on main window event
    if(event.widget.master == None):
        # width and height of Ledger_Viewer
        RVwidth = event.width
        RVheight = event.height
        header = 50

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
        my_tree.column("Student ID", anchor=W, width= 4*CW)
        my_tree.column("Student", anchor=W, width= 30*CW)
        my_tree.column("Registered", anchor=W, width= CW)
        my_tree.column("Start Term", anchor=W, width= CW)
        my_tree.column("Grade", anchor=W, width= 8*CW)
        my_tree.column("Major Type", anchor=W, width= CW)
        my_tree.column("Major", anchor=W, width= 20*CW)
        my_tree.column("Cl", anchor=W, width= 8*CW)
        my_tree.column("Email", anchor=W, width= 60*CW)
        my_tree.column("Registration Status", anchor=W, width= CW)
        my_tree.place(x = 10, y = header, width = RVwidth-30, height = RVheight-header-10)

        # change the location and size of the scrollbar
        vsb.place(x=RVwidth-20, y=header, height=RVheight-header-10)  

def print_roster(event):
    k = 0
    first_row_in_class = 0
    even_odd = 0
    prev_class = ''

    #clear display
    my_tree.delete(*my_tree.get_children())

    #check for populated department
    students = []
    if(drop_dept.get() != ""):
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
        

    for row in complete_roster:
        print_roster = False
        if(drop_dept.get() == '' and drop_student.get() == ''):
            print_roster = True
        elif(drop_dept.get() == row[1] and drop_student.get() == ''):
            print_roster = True
        elif(drop_student.get() == row[13]):
            print_roster = True

        if(print_roster == True):
            # get the course number
            num = re.findall(r'\d+', row[2])
            if(((locale.atof(row[11]) < 10) and (locale.atof(num[0]) < 500)) or (locale.atof(row[11]) < 6)):
                color = 'negative'
            else:
                color = 'positive'
            # if students are in the same class
            if(row[2] == prev_class):
                even_odd -= 1
            if even_odd % 2 == 0:
                my_tree.insert(parent='', index='end', iid=k, text="", values=row, tags=('evenrow',color))
            else:
                my_tree.insert(parent='', index='end', iid=k, text="", values=row, tags=('oddrow',color))

            # if students are in the same class
            if(row[2] == prev_class):
                my_tree.move(k, first_row_in_class, 0)
            else:
                first_row_in_class = k
            even_odd += 1

            k+=1
            prev_class = row[2]

    
# define dept combo box
drop_dept = ttk.Combobox(root,values = "", state="readonly", width=25)
dept_label = Label(root, text="Dept.")
dept_label.place(x=100, y=10)
drop_dept.place(x=150, y=10)
drop_dept.bind("<<ComboboxSelected>>", print_roster)

# define student combo box
drop_student = ttk.Combobox(root,values = "", state="readonly", width=25)
student_label = Label(root, text="Student")
student_label.place(x=350, y=10)
drop_student.place(x=400, y=10)
drop_student.bind("<<ComboboxSelected>>", print_roster)

#bind GUI resizing
root.bind('<Configure>', resizer)

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

            # fill in the student dropdown
            students = []
            for listrow in matrix:
                row = list(map(str,listrow))
                try:
                    if(row[0].isnumeric()):
                        original = True
                        for i in students:
                            if(i == row[13]):
                                original = False
                                break
                        if(original == True):
                            students.append(row[13])
                except:
                    continue
            students.sort()
            students.insert(0,"")
            drop_student.set('')
            drop_student['values'] = students

            # show the name of the file
            root.title(getfile)
 
            print_roster(0)

# open button
open_button = ttk.Button(
    root,
    text='Open File',
    command=select_file
)
# put it in the upper left corner
open_button.place(x=10, y=5)

import locale 

my_tree.place(x = 10, y = header, width = RVwidth-30, height = RVheight-header-10)

root.mainloop()