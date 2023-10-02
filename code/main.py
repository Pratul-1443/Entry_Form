import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os


def Enter():
    temp = accept.get()
    if temp == "Accepted":

        # for user
        title = title_combobox.get()
        first_name = first_Name_entry.get()
        last_name = last_Name_label.get()
        age = age_spin.get()
        nationality = nationality_combobox.get()
        status = status_list.get()
        # for Courses
        check = ref_status_var.get()
        higher = higher_couse.get()
        year = year_entry.get()
        if first_name and last_name and nationality and higher:
            print(f"Title : {title} First Name : {first_name}  Last Name : {last_name}")
            print(f"Age :{age}, Nationality : {nationality} Status :{status}")
            print(f"Higher Eduction: {higher} Year : {year}")
            print(f"Registration Status :{check} ")

            filepath = "C:/Users/Pratul/OneDrive/Desktop/PY_Data/data.xlsx"
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["Title", "First Name", "Last Name", "Age", "Nationality", "Status", "Higher Education",
                           "Registration status", "Year"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook=openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([title, first_name, last_name, age, nationality, status, higher,
                       check, year])
            workbook.save(filepath)

        else:
            tkinter.messagebox.showwarning(title="Error", message="Entry Not Fill")
    else:
        tkinter.messagebox.showwarning(title="Error", message="Check term and Conditions")
    print("-------------------------------------------------")


root = Tk()
root.title("Data Entry Form")
frame = Frame(root,background="#36454F")
frame.pack()

# 1

user = LabelFrame(frame, text="User Info",background="#899499")
user.grid(row=0, column=0, padx=20, pady=10)

title = Label(user, text="Title")
title_combobox = ttk.Combobox(user, values=["", "Mr.", "Ms.", "Dr."])
title.grid(row=0, column=0)
title_combobox.grid(row=1, column=0)

first_Name_label = Label(user, text="First Name")
first_Name_label.grid(row=0, column=1)

last_Name_label = Label(user, text="Last Name")
last_Name_label.grid(row=0, column=2)

first_Name_entry = Entry(user)
last_Name_label = Entry(user)

first_Name_entry.grid(row=1, column=1)
last_Name_label.grid(row=1, column=2)

age_label = Label(user, text="Age")
age_spin = Spinbox(user, from_=17, to=100)
age_label.grid(row=2, column=0)
age_spin.grid(row=3, column=0)

nationality_label = Label(user, text="Nationality")
nationality_combobox = ttk.Combobox(user, values=["None", "Bharat", "USA", "Japan", "Australia"])

nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

status = Label(user, text="Status")
status_list = ttk.Combobox(user, values=["None", "Single", "Married"])
status.grid(row=2, column=2)
status_list.grid(row=3, column=2)

for i in user.winfo_children():
    i.grid_configure(padx=10, pady=5)

# 2 box

course = LabelFrame(frame, text="Courses",background="#899499")
course.grid(row=1, column=0, padx=20, pady=10, sticky="news")


higher_edu = Label(course, text="Higher Education")
higher_couse = ttk.Combobox(course, values=["None", 10, 12, "Graduand", "Graduate"])
higher_edu.grid(row=0, column=0)
higher_couse.grid(row=1, column=0)

year = Label(course, text="Year")
year_entry = Entry(course)
year.grid(row=0, column=2)
year_entry.grid(row=1, column=2)

regiser = Label(course, text="Registration Status")

ref_status_var = StringVar(value="NotRegistered")
regiser_chcek = Checkbutton(course, text="Currently Registered", variable=ref_status_var, onvalue="Register",
                            offvalue="Not Register")
regiser.grid(row=0, column=1)
regiser_chcek.grid(row=1, column=1)

for i in course.winfo_children():
    i.grid_configure(padx=10, pady=5)

# 3 box

term_condition = LabelFrame(frame, text="Term & Conditions",background="#899499")
term_condition.grid(row=2, column=0, sticky="news", pady=10, padx=20)

accept = StringVar(value="not accepted")
term_check = Checkbutton(term_condition, text="I accept the terms and conditions", variable=accept, onvalue="Accepted",
                         offvalue="Not Accepted")
term_check.grid(row=0, column=0)

# 4

button =Button(
    frame,
    text='Submit',
    font=('Times', 26),
    padx=10,
    pady=10,
    bg='#899499',
    fg='white',
    command=Enter,
    activebackground='green',
    activeforeground='red'
    )
button.grid(row=3, column=0, sticky="news", pady=10, padx=20)

root.mainloop()
