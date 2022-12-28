#importing tkinter modules
import os
from tkinter import *
from tkinter import messagebox
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import openpyxl
import statistics



# creating a new tkinter window
root = Tk()

# assigning a title
root.title("Student Evaluation System")

# specifying geometry for window size
root.geometry("800x650")

frame = LabelFrame(root,text="STUDENT EVALUATION SYSTEM",padx=30,pady=30)
frame.pack(padx=10,pady=10)

# ---------- FUNCTIONS ----------
def quit_window():
    root.destroy()

def results_all():
    get_result()
    create_linegraph()
    create_bargraph()

def submit_all():
    get_grades()
    popup()
    insert_excel()

def get_result():
    dfr = pd.read_excel('data.xlsx')
    subjects_total = dfr['Subject1']
    print(subjects_total)

def create_linegraph():
    dfr = pd.read_excel('data.xlsx')
    subjectlist = [dfr['Subject1'],dfr['Subject2'],dfr['Subject3'],dfr['Subject4'],dfr['Subject5']]
    for subject in subjectlist:
        mean = statistics.mean(subject)
        stddev = statistics.stdev(subject)
        plotdf = pd.DataFrame({
            'data': subject,
            'mean': [mean for i in range(len(subject))],
            'std': [stddev for i in range(len(subject))]
        })

        plotdf.plot()
        plt.title("Report")
        plt.show()

def create_bargraph():
    dfr = pd.read_excel('data.xlsx')
    X = ['Subject1','Subject2','Subject3','SUbject4', 'Subject5']
    maximum_marks = [
        dfr['Subject1'].max(),
        dfr['Subject2'].max(),
        dfr['Subject3'].max(),
        dfr['Subject4'].max(),
        dfr['Subject5'].max(),
        ]
    minimum_marks = [
        dfr['Subject1'].min(),
        dfr['Subject2'].min(),
        dfr['Subject3'].min(),
        dfr['Subject4'].min(),
        dfr['Subject5'].min(),
    ]
    you_marks = [
        int(e_s1.get()),
        int(e_s2.get()),
        int(e_s3.get()),
        int(e_s4.get()),
        int(e_s5.get())
        ]

    X_axis = np.arange(len(X))
    
    plt.bar(X_axis - 0.2, maximum_marks, 0.2, label = 'Max-marks')
    plt.bar(X_axis + 0.2, minimum_marks, 0.2, label = 'Min-marks')
    plt.bar(X_axis, you_marks, 0.2, label = 'Your-marks')
    
    plt.xticks(X_axis, X)
    plt.xlabel("Subjects")
    plt.ylabel("Marks")
    plt.title("Report")
    plt.legend(loc='upper left')
    plt.show()


def insert_excel():
    df1 = pd.read_excel('data.xlsx')
    sum = int(e_s1.get()) + int(e_s2.get()) +int(e_s3.get())+int(e_s4.get())+int(e_s5.get())
    per = sum/5.0
    df2 = pd.DataFrame([[e_name.get(),e_uid.get(),e_s1.get(),e_s2.get(),e_s3.get(),e_s4.get(),e_s5.get(),sum,per]],columns=['Name','UID','Subject1','Subject2','Subject3','Subject4','Subject5','Total','Percentage'])
    df3 = pd.concat([df1,df2])
    df3.to_excel('data.xlsx',sheet_name='Marks',index=FALSE)

def popup():
    messagebox.showinfo("Message","Data Saved")

def get_grades():
    marks = [e_s1.get(),e_s2.get(),e_s3.get(),e_s4.get(),e_s5.get()]
    for i in range(6,11):
        if(int(marks[i-6])>100):
            Label(frame, text="ERR").grid(row=i, column=3,padx=50,pady=10)
        elif(int(marks[i-6])>95):
            Label(frame, text="A+").grid(row=i, column=3,padx=50,pady=10)
        elif(int(marks[i-6])>90):
            Label(frame, text="A").grid(row=i, column=3,padx=50,pady=10)
        elif(int(marks[i-6])>80): 
            Label(frame, text="B").grid(row=i, column=3,padx=50,pady=10)
        elif(int(marks[i-6])>70):
            Label(frame, text="C").grid(row=i, column=3,padx=50,pady=10)
        elif(int(marks[i-6])>60):
            Label(frame, text="D").grid(row=i, column=3,padx=50,pady=10)
        else:
            Label(frame, text="E").grid(row=i, column=3,padx=50,pady=10)

    get_results = Button(frame,width=15,padx=10,pady=10,text="Get Results",command=results_all).grid(row=12,column=4)

# ---------- IMPORT IMAGES ----------
quit_img = PhotoImage(file="Power-Off-Logo.png")


# ---------- WIDGETS -----------
btn_quit = Button(frame, image=quit_img, command=quit_window,borderwidth=0,width=30,height=30)
btn_quit.place(relx = 1, x =-2, y = 2,anchor=NE)

Label(frame,text="       ").grid(row=0,column=0)
Label(frame,text="       ").grid(row=3,column=0)
Label(frame,text="       ").grid(row=11,column=0)
Label(frame,text="\t\t\t\t").grid(row=11,column=4)

Label(frame,text="Student's Name:     ").grid(row=1,column=1,padx=10,pady=10)

e_name = Entry(frame,width=18,borderwidth=1)
e_name.grid(row=1,column=2)


Label(frame,text="Student's UID No:     ").grid(row=2,column=1,padx=10,pady=10)

e_uid = Entry(frame,width=18,borderwidth=1)
e_uid.grid(row=2,column=2)

Label(frame,text="Enter Marks to Analyze:     ").grid(row=4,column=1,padx=8,pady=8)

# SECTION FOR MARKS AND GRADES
Label(frame, text="SR.No").grid(row=5, column=1,padx=10,pady=10)
Label(frame, text="Marks").grid(row=5, column=2,padx=10,pady=10)
Label(frame, text="Grades").grid(row=5, column=3,padx=50,pady=10)
Label(frame, text="1").grid(row=6, column=1,padx=10,pady=10)
Label(frame, text="2").grid(row=7, column=1,padx=10,pady=10)
Label(frame, text="3").grid(row=8, column=1,padx=10,pady=10)
Label(frame, text="4").grid(row=9, column=1,padx=10,pady=10)
Label(frame, text="5").grid(row=10, column=1,padx=10,pady=10)

e_s1 = Entry(frame,width=20)
e_s2 = Entry(frame,width=20)
e_s3 = Entry(frame,width=20)
e_s4 = Entry(frame,width=20)
e_s5 = Entry(frame,width=20)

e_s1.grid(row=6, column=2,padx=10,pady=10)
e_s2.grid(row=7, column=2,padx=10,pady=10)
e_s3.grid(row=8, column=2,padx=10,pady=10)
e_s4.grid(row=9, column=2,padx=10,pady=10)
e_s5.grid(row=10, column=2,padx=10,pady=10)


btn_submit = Button(frame, text="Submit",command=submit_all, width=15, padx=10, pady=10).grid(row=12,column=3)

# Rendering the main window
frame.mainloop()
