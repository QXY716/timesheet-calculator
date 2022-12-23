import xlrd
import xlwt
import tkinter as tk
import os
import sys

# filepath = os.path.join(os.path.dirname(os.path.realpath(sys.executable)),'timesheet_input.xls')
filepath = './timesheet_input.xls'

def create_table():
    workbook = xlwt.Workbook(encoding= 'ascii')
    worksheet = workbook.add_sheet("Sheet1")
    worksheet.write(0,0, "Name")
    worksheet.write(0,1, "Department/Proffesion")
    for i in range(31):
        worksheet.write(0,2+i,str(i+1)+' day')
    worksheet.write(0,33, "Days/hours from 1 to 15")
    worksheet.write(0,34, "Day/hours from 1 to 31")
    workbook.save("timesheet_input.xls")

part_time = {'Ш_Тр':0.25,'Щ_Тр ':0.25,'Н_Тр':0.25,'К_Тр':0.2,'М_Тр':0.1,'С_О':0.25,'Ш_1_О':0.25,'Ш_2_О':0.25}
department = {'Ш_Тр':'Tr','Щ_Тр ':'Tr','Н_Тр':'Tr','К_Тр':'Tr','М_Тр':'Tr','С_О':'o','Ш_1_О':'o','Ш_2_О':'o'}

data = xlrd.open_workbook(filepath)
datasheet = data.sheets()[0]
nrow = datasheet.nrows
ncol = datasheet.ncols

def get_result(row):
    name = datasheet.cell_value(row,0)
    pos = datasheet.cell_value(row,1)
    rowdata=datasheet.row_values(row,2,33)
    rate = part_time[pos]
    dep = department[pos]
    day15 = 0
    day31 = 0
    for i in range(15):
        if rowdata[i]=='О':
            day15+=1
    for j in range(31):
        if rowdata[j]=='О':
            day31+=1
    result = ('Employee\'s name: ' + str(name) + '\n'
              'Department: ' + str(dep) + '\n'
              'In 15 days: ' + str(day15) + ' Days' + '/' + str(rate*day15*8) + ' hours' + '\n'
              'In a month: ' + str(day31) + ' Days' + '/' + str(rate*day31*8) + ' hours')
    return result

def show_result(name):
    namelist = datasheet.col_values(0)
    try:
        index = namelist.index(name)
        text = get_result(index)
        return text
    except:
        error = 'No employee found'
        return error

window=tk.Tk()
window.title('timesheet calculator')
window.geometry('800x640')
text = tk.StringVar()
entry = tk.Entry(window,textvariable=text)
title = tk.Label(window,text='Enter the employee\'s name',bg='white',font=('Arial', 12))
title.place(relx=0.5,rely=0.1,anchor="center")
text.set('')
entry.place(relx=0.5,rely=0.2,anchor="center",height=40,width=200)
t = tk.Text(width=50, height=25)
t.place(relx=0.5,rely=0.7,anchor="center")
def printEntry():
    var = show_result(text.get())
    t.delete("1.0","end")
    t.insert('end', var)
button = tk.Button(window,text='calculate',command=printEntry,height=2,width=20,font=('Arial', 12))
button.place(relx=0.5,rely=0.35,anchor="center")

def main():
    window.mainloop()

if __name__ == '__main__':
    try:
        if sys.argv[1] == 'create':
            create_table()
        elif sys.argv[1] == 'calculate':
            main()
        else:
            print('The parameter is wrong, please enter \'create\'to create the timesheet')
    except:
        main()
