# STUDENT-FORM-with-excel-intergrated
#It is basic python project containing GUI which have small feature but not connected with data base now 


import tkinter,openpyxl
from openpyxl import Workbook,load_workbook
from tkinter import Button,Entry,ttk, messagebox
from openpyxl.utils import get_column_letter
run=tkinter.Tk()
run.geometry("1920x1080")

run.title("STUDENT INTRODUCTION")
run.configure(bg="pink")

run.resizable(True,True)
#for excel integration
Work_book = openpyxl.load_workbook("Book1.xlsx")
work_sheet = Work_book.active
if work_sheet["A1"].value is None:
    work_sheet["A1"] = "NAME"
    work_sheet["B1"] = "COLLEGE"
    work_sheet["C1"] = "PROGRAM"
    work_sheet["D1"] = "SUBJECT"
    work_sheet["E1"] = "GOAL"
    work_sheet["F1"] = "SEMESTER"
    Work_book.save("Book1.xlsx")

    #to  clear the information that has been submited
def onclick():
    name = Enter_name1.get()
    college = college_name1.get()
    program_name = program1.get()
    subject = favorite_subject1.get()
    semester = dropdown.get()
    goal = carer_goal1.get()
    if name and college and program_name and subject and goal and semester:
        try:
           #for saving information in excel
           work_sheet.append([name,college,program_name,subject,goal,semester])
           Work_book.save("Book1.xlsx")
           #this delete all the data after submitting the data
           Enter_name1.delete(0,tkinter.END)
           college_name1.delete(0,tkinter.END)
           program1.delete(0,tkinter.END)
           favorite_subject1.delete(0,tkinter.END)
           carer_goal1.delete(0,tkinter.END)
           dropdown.set("")
           messagebox.showinfo("SUCCESS","YOUR DATA HAS BEEN SUBMITTED")

        
        except PermissionError:
         #so to make sure that the excel file is closed before submitting the data
         messagebox.showerror("ERROR","PLEASE CLOSE THE EXCEL FILE BEFORE SUBMITTING")
    else:
           #to make sure that all the information is filled before submitting the data
           tkinter.messagebox.showerror("ERROR","PLEASE FILL ALL THE INFORMATION")
   





#for entering name

Enter_name=tkinter.Label(run,text="ENTER YOUR NAME")
Enter_name.configure(font=("Arial",12,"bold"),fg="black",bg="white")
Enter_name.pack()
#space box for entering name

Enter_name1 = tkinter.Entry(run)
Enter_name1.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
Enter_name1.pack()

#for college name
college_name = tkinter.Label(run,text="ENTER COLLEGE NAME")
college_name.configure(font=("Arial",12,"bold"),fg="black",bg="white")
college_name.pack()

#for entry space of college name
college_name1 = tkinter.Entry(run)
college_name1.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
college_name1.pack()

#for program entry space and dispaly text
program = tkinter.Label(run,text= "ENTER PROGRAM NAME")
program.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
program.pack()
program1 = tkinter.Entry(run)
program1.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
program1.pack()

#for favorite subject

favorite_subject= tkinter.Label(run,text= "ENTER FAVORITE SUBJECT")
favorite_subject.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
favorite_subject.pack()
favorite_subject1 =tkinter.Entry(run)
favorite_subject1.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
favorite_subject1.pack()


#for carer goal

carer_goal = tkinter.Label(run,text="ENTER CAREER GOAL")
carer_goal.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
carer_goal.pack()
carer_goal1 = tkinter.Entry(run)
carer_goal1.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
carer_goal1.pack()

#for dropdown
choice = [1,2,3,4,5,6,7,8]
semester = tkinter.Label(run,text="ENTER YOUR SEMESTER")
semester.configure(font=("Arial",12,"bold"),fg="black",bg="white",width="30")
semester.pack()
dropdown = ttk.Combobox(run,values=choice)
dropdown.pack()

#for selection
select = tkinter.Button(run,text= "SUBMIT ",command=onclick)
select.configure(font=("Arial",12,"bold"),fg="black",bg="white",width=15,padx=40,pady=10)
select.pack(anchor="center")


run.mainloop()

