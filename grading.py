#!/usr/bin/env python               #solving the Turkish character problem
#-*- coding: utf-8 -*-

from Tkinter import *
import tkMessageBox
import tkFileDialog
import xlrd
import os
import sys
import anydbm
reload(sys)
sys.setdefaultencoding('utf-8')



class Grading_file():
    def __init__(self):
        self.line=[]
    def open_file(self):
        student_info= {}
        file = tkFileDialog.askopenfilename(initialdir=os.getcwd(), title="Please Select a File",filetypes=[('excel files', ('.xlsx', '.xls'))])
        worbook = xlrd.open_workbook(file, "rb")  # reading excel file's informations
        sheet = worbook.sheet_by_index(0)
        self.line = [str(sheet.cell_value(0, 2)),
                     str(sheet.cell_value(0, 3))]
        list1 = []
        row = 1
        while row <= sheet.nrows - 1:
            if student_info.has_key(sheet.cell_value(row, 2)) == True:  # check the section are in dict or not
                info_dict = {str(sheet.cell_value(row, 0)): (
                str(sheet.cell_value(row, 2)), str(sheet.cell_value(row, 3)))}
                student_info[str(sheet.cell_value(row, 2))].append(info_dict)
            else:
                info_dict = {str(sheet.cell_value(row, 0)): (str(sheet.cell_value(row, 1)), str(sheet.cell_value(row, 2)))}
                list1.append(info_dict)
                student_info[str(sheet.cell_value(row, 3))] = list1

            list1 = []
            row += 1

        return student_info

class Grade_Calculator(Frame):

    def __init__(self,parent):
        self.parent = parent
        Frame.__init__(self,parent)
        self.initUI(parent)


    def initUI(self,parent):
        self.headlabel = Label(self, text="ENGR102 Numerical Grade Calculator", font=("", "15", "bold"))  # headline of the app
        self.headlabel.grid(row=0, column=1,columnspan=7)

        self.select_std_label = Label(self, text="MP1 %", font=("", "12"))
        self.select_std_label.grid(row=1, column=0)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry1 = Entry(self, width=6,textvariable=self.variable)  # It creates an entry widget and we will give some value to calculate attendance.
        self.threshold_entry1.grid(row=1, column=1,padx=10)

        self.select_std_label = Label(self, text="MP2 %", font=("", "12"))
        self.select_std_label.grid(row=1, column=2,ipadx=10)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry2 = Entry(self, width=6,textvariable=self.variable)
        self.threshold_entry2.grid(row=1, column=3)

        self.select_std_label = Label(self, text="MP3 %", font=("", "12"))
        self.select_std_label.grid(row=1, column=4,ipadx=20)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry3 = Entry(self, width=6, textvariable=self.variable)
        self.threshold_entry3.grid(row=1, column=5)

        self.select_std_label = Label(self, text="MP4 %", font=("", "12"))
        self.select_std_label.grid(row=1, column=6,ipadx=20)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry4 = Entry(self, width=6, textvariable=self.variable)
        self.threshold_entry4.grid(row=1, column=7)

        self.select_std_label = Label(self, text="MP5 %", font=("", "12"))
        self.select_std_label.grid(row=1, column=8,ipadx=15)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry5 = Entry(self, width=6, textvariable=self.variable)
        self.threshold_entry5.grid(row=1, column=9)

        self.select_std_label = Label(self, text="Midterm %", font=("", "12"))
        self.select_std_label.grid(row=2, column=0)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry6 = Entry(self, width=6, textvariable=self.variable)
        self.threshold_entry6.grid(row=2, column=1)

        self.select_std_label = Label(self, text="Final %", font=("", "12"))
        self.select_std_label.grid(row=3, column=0)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry7 = Entry(self, width=6, textvariable=self.variable)
        self.threshold_entry7.grid(row=3, column=1)

        self.select_std_label = Label(self, text="Attendance %", font=("", "12"))
        self.select_std_label.grid(row=4, column=0)

        self.variable = StringVar()  # It helps us the set variable.
        self.threshold_entry8 = Entry(self, width=6, textvariable=self.variable)
        self.threshold_entry8.grid(row=4, column=1)

        self.select_std_label = Label(self, text="Grading File %", font=("", "12"))
        self.select_std_label.grid(row=3, column=2,columnspan=3)

        self.select_std_label = Label(self, text="Attendance File %", font=("", "12"))
        self.select_std_label.grid(row=4, column=2,columnspan=3)

        self.label_calculate=Label(self)
        self.label_calculate.grid(row=5,column=2)

        self.label_x=Label(self)
        self.label_x.grid(row=7,column=10)


        self.import_button = Button(self, text="Calculate", font=("", "8"),command=self.calculating_grades)
        self.import_button.grid(row=6, column=2)

        self.save_button=Button(self,text='Save',font=('','9'),command=self.savingfile)
        self.save_button.grid(row=6,column=4)

        self.browse_button = Button(self, text='Browse', width=10, bg='red',command=self.browsing_gradingfile)
        self.browse_button.grid(row=3, column=5)
        self.browse_button.config(relief=RIDGE)

        self.browse_button = Button(self, text='Browse', width=10, bg='red',command=self.browsing_attendancefile)
        self.browse_button.grid(row=4, column=5)
        self.browse_button.config(relief=RIDGE)

        scrollbar = Scrollbar(self, orient=VERTICAL)
        self.student_list = Listbox(self, selectmode="multiple",height=19,width=110,yscrollcommand=scrollbar.set)
        self.student_list.grid(row=8,column=0,rowspan=7,columnspan=10,padx=18)
        scrollbar.config(command=self.student_list.yview)
        scrollbar.grid(row=8,column=9,rowspan=7,sticky=N+S)

        self.reload_db()



    def reload_db(self):
        if os.path.isfile("./grades.db"):
            self.db = anydbm.open("grades.db", "r")
            list=[]
            for k,v in self.db.iteritems():
                list.append(k)
            list.sort()
            for i in list:
                self.student_list.insert(END,(" "*15 + i + "    " +str(self.db[i])[:6]))
            self.db.close()

    def browsing_gradingfile(self):

        self.student_dict = {}  # dict stor the class and student infos
        file = tkFileDialog.askopenfilename(initialdir=os.getcwd(), title="Please Select a File",
                                            # opening the excel file which is using in our app
                                            filetypes=[('excel files', ('.xlsx', '.xls'))])

        worbook = xlrd.open_workbook(file, "rb")  # reading excel file's informations
        sheet = worbook.sheet_by_index(0)
        row = 1
        while row <= sheet.nrows-1:
            self.student_dict[str(sheet.cell_value(row,2))+" "+str(sheet.cell_value(row,3))] = [str(sheet.cell_value(row,6)),
                                                            str(sheet.cell_value(row,7)),str(sheet.cell_value(row,8)),
                                                            str(sheet.cell_value(row,9)),str(sheet.cell_value(row,10)),
                                                            str(sheet.cell_value(row,11)),str(sheet.cell_value(row,12))]
            row += 1






    def browsing_attendancefile(self):
        self.attendence_dict = {}  # dict stor the class and student infos
        file = tkFileDialog.askopenfilename(initialdir=os.getcwd(), title="Please Select a File",
                                            # opening the excel file which is using in our app
                                            filetypes=[('excel files', ('.xlsx', '.xls'))])

        worbook = xlrd.open_workbook(file, "rb")  # reading excel file's informations
        sheet = worbook.sheet_by_index(0)
        row = 2
        while row <= sheet.nrows - 1:
            self.attendence_dict[str(sheet.cell_value(row,0))+" "+str(sheet.cell_value(row,1))] = 0
            column = 3
            while column <= 16:
                if sheet.cell_value(row,column) == 1:
                    self.attendence_dict[str(sheet.cell_value(row, 0)) + " " + str(sheet.cell_value(row, 1))]+= 1
                column += 1
            row += 1


        sheet2 = worbook.sheet_by_index(1)
        row = 2
        while row <= sheet2.nrows - 1:
            column = 3
            while column <= 16:
                if sheet2.cell_value(row,column) == 1:
                    self.attendence_dict[str(sheet2.cell_value(row, 0)) + " " + str(sheet2.cell_value(row, 1))] += 1
                column += 1
            row += 1




    def savingfile(self):
        self.db = anydbm.open("grades.db", "n")
        for i in self.list:
            self.db[i] = str(self.calculating_dict[i])
        self.db.close()

    def calculating_grades(self):
        self.student_list.delete(0,END)
        entry_total = int(self.threshold_entry1.get())+int(self.threshold_entry2.get())+int(self.threshold_entry3.get())+\
                      int(self.threshold_entry4.get())+int(self.threshold_entry5.get())+int(self.threshold_entry6.get())+\
                      int(self.threshold_entry7.get())+int(self.threshold_entry8.get())
        if entry_total == 100:
            self.list = self.student_dict.keys()
            self.list.sort()
            self.calculating_dict={}
            for i in self.list:
                self.calculating_dict[i]=((float(self.student_dict[i][0])*int(self.threshold_entry1.get())/100)
                                          +(float(self.student_dict[i][1])*int(self.threshold_entry2.get())/100)+
                (float(self.student_dict[i][2]) * int(self.threshold_entry3.get()) / 100)+
                                          (float(self.student_dict[i][3])*int(self.threshold_entry4.get())/100)+
                (float(self.student_dict[i][4]) * int(self.threshold_entry5.get()) / 100)+
                                          (float(self.student_dict[i][5])*int(self.threshold_entry6.get())/100)+
                (float(self.student_dict[i][6]) * int(self.threshold_entry7.get()) / 100)+
                                          ((float(self.attendence_dict[i])*100/28) * int(self.threshold_entry8.get()) / 100))
                std = " "*15 + i + "    " + str(self.calculating_dict[i])[:6]
                self.student_list.insert(END, std)

        else:
            tkMessageBox.showinfo("Warning!!!!", "The assessment components do NOT sum up to 100")


def main():
    root = Tk()
    root.title('Grade Calculator')
    root.geometry('750x520')
    app = Grade_Calculator(root)
    app.pack(fill=BOTH, expand=True)
    root.mainloop()
main()
