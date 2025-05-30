import openpyxl
import os
from natsort import os_sorted 
from tkinter import *
from logging import *

# excel_file = 'reports.xlsx'
# wb = openpyxl.load_workbook(excel_file)
# sheet = wb['Sheet1']
rev_letters = ['A','B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N']




class App:
    def __init__(self, master):

        self.master = master 

        # Data vars
        self.report_number = StringVar()
        self.report_title = StringVar()
        self.report_type = StringVar()
        self.system = StringVar()
        self.revisions = StringVar()
        self.add_nc = StringVar()
        self.add_labels = []
        self.add_entries = []
        self.add_variables = []

        # Title and geometry
        master.title("DocFinder")
        master.geometry('700x700')
        master.resizable(True, True)

        # Entries and Labels

        # Report Numbers, Name and Type
        self.report_num_label = Label(master, text="Report Num")
        self.report_num_label.grid(row = 0, column = 0)
        self.report_num_entry = Entry(master, textvariable=self.report_number)
        self.report_num_entry.grid(column =1 , row = 0, padx=5, pady = 5)
        
        self.report_title_label = Label(master, text="Report Title")
        self.report_title_label.grid(row = 1, column = 0)
        self.report_title_entry = Entry(master, textvariable=self.report_title)
        self.report_title_entry.grid(column =1 , row = 1, padx=5, pady = 5)

        self.report_type_label = Label(master, text="Report Type")
        self.report_type_label.grid(row = 2, column = 0)
        self.report_type_entry = Entry(master, textvariable=self.report_type)
        self.report_type_entry.grid(column =1 , row = 2, padx=5, pady = 5)


        # System 
        self.system_label = Label(master, text="System")
        self.system_label.grid(row = 3, column = 0)
        self.system_entry = Entry(master, textvariable = self.system)
        self.system_entry.grid(column =1 , row = 3, padx=5, pady = 5)

        # Number of revisions
        self.rev_amount_label = Label(master, text="Number of Revs")
        self.rev_amount_label.grid(column = 0, row = 4)
        self.rev_amount_entry = Entry(master, textvariable=self.revisions)
        self.rev_amount_entry.grid(column = 1, row = 4, padx=5, pady = 5)

        # Addendums for Rev -- 
        self.add_nc_label = Label(master, text="Addendum Num for Rev --" )
        self.add_nc_label.grid(column = 0, row = 5)
        self.add_nc_entry =  Entry(master, textvariable=self.add_nc)
        self.add_nc_entry.grid(column = 1, row = 5)
        

        # Traces
        self.revisions.trace_add("write", callback=self.revisions_updated_trace)

        # Submit Button
        self.submit_btn = Button(master, text="Send Data to Excel", fg="white", bg = "Red", command = self.submit)
        self.submit_btn.grid(column = 0, row = 6, padx=5, pady =5)
    
    def revisions_updated_trace(self,*args): # Since revisions are dynamic
        self.revisions.get()
        if self.revisions.get() != "":
            if int(self.revisions.get()) > 0:
                for i in range(int(self.revisions.get())):
                    self.add_variables.append(StringVar())
                    self.add_labels.append(Label(self.master, text = f"Addendum Num for Rev {rev_letters[i]}"))
                    self.add_entries.append(Entry(self.master,textvariable=self.add_variables[i]))
                    self.add_labels[i].grid(row = i+6, column = 0, padx=5, pady=5)
                    self.add_entries[i].grid(row = i+6, column = 1, padx=5, pady = 5)
                    self.submit_btn.grid(column = 0, row = 7 + i,padx=5,pady=5)
        elif self.revisions.get() == "" and len(self.add_labels) > 0:   
            for i in range(len(self.add_labels)):
                self.add_labels[i].destroy()
                self.add_entries[i].destroy()
            self.add_labels.clear()
            self.add_entries.clear()
            self.submit_btn.grid(column = 0, row = 6,padx=5, pady=5)
    
    def submit(self): # PyXl code 
        print(f"{self.report_number.get()} --") # Default revision
        if int(self.add_nc.get()) > 0:
            for i in range(int(self.add_nc.get())+1):
                if i != 0:
                    print(f"{self.report_number.get()}---{i}")
        for i in range(int(self.revisions.get())):
            for j in range(int(self.add_variables[i].get())+1):
                if j != 0:
                    print(f"{self.report_number.get()} {rev_letters[i]}-{j}")
                else: 
                    print(f"{self.report_number.get()} {rev_letters[i]}")
        self.master.destroy()
        
            





root = Tk()
app_instance = App(root)
root.mainloop()