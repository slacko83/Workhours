import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import calendar
import datetime
import time
import tkinter as tk
from tkinter import messagebox

#TODO: hibakezelés(inputok formátuma, megnyitott excel esetén, stb), visszajelzés a felvitt dolgozókrol,
# évszám legyen protected???, conditional formatting

"""Create class for the employees"""
class Employee:
    def __init__(self, name, workstart1, workend1, workstart2, workend2, holidays):
        self.name = name
        self.workstart1 = workstart1
        self.workend1 = workend1
        self.workstart2 = workstart2
        self.workend2 = workend2
        self.holidays = holidays

def create_excel(employees, year):
    """Set variables"""
    sheet = []
    months_hu = {'January': 'Január', 'February': 'Február', 'March': 'Március', 'April': 'Április', 'May': 'Május',
                'June': 'Június', 'July': 'Július', 'August': 'Augusztus', 'September': 'Szeptember', 'October': 'Október',
                'November': 'November', 'December': 'December'}
    days_hu={'vasárnap':6, 'hétfő':0, 'kedd':1, 'szerda':2,'csütörtök':3, 'péntek':4,'szombat':5}

    """Create the workbook and set the formats used in the excel"""
    workbook = xlsxwriter.Workbook('sample_data4.xlsx')
    title_format = workbook.add_format({"bold": True, "font_name": "Segoe UI", "font_size": 12})
    header_format = workbook.add_format({"bold": True, "font_name": "Segoe UI", "font_size": 8})
    year_format = workbook.add_format({"bold": True, "font_name": "Segoe UI", "font_size": 12, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":2, "left":2, "right":2, "bottom":1})
    format = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter"})
    format_name = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":2, "left":2, "right":2, "bottom":1})
    format_left = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color": "#C0C0C0", "top": 1, "left": 2, "right": 1, "bottom": 1})
    format_right = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":1, "right":2, "bottom":1})
    format_whole = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "border":2})
    date_format = workbook.add_format({'num_format': '[$-40E]mmmm d\.;@', "font_name": "Segoe UI", "font_size": 8, "align": "center", "border_color":"#C0C0C0", "top":1, "left":2, "right":1, "bottom":1})
    day_format = workbook.add_format({'num_format': 'dddd', "font_name": "Segoe UI", "font_size": 8, "align": "center", "border_color":"#C0C0C0", "top":1, "left":1, "right":2, "bottom":1})
    time_format = workbook.add_format({'num_format': 'h:mm', "font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "border":1})
    time_format_left = workbook.add_format({'num_format': 'h:mm', "font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":2, "right":1, "bottom":1})
    worktime_format = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":1, "right":2, "bottom":1})
    holiday_format = workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "bg_color":"#C0C0C0"})
    legend_format = workbook.add_format({"font_name": "Segoe UI", "font_size": 8})

    """Cycle of worksheets"""
    for i in range(0, 12):
        """Create the worksheets and set the size of rows and columns"""
        sheet.append(workbook.add_worksheet(months_hu[calendar.month_name[i + 1]]))
        sheet[i].set_column("A:Z", 5.57)
        sheet[i].set_column("A:A", 11)
        sheet[i].set_column("B:B", 9.14)

        """Fill the  constant fields"""
        sheet[i].write('A1', 'MUNKAIDŐ NYILVÁNTARTÁS', title_format)
        sheet[i].write('F1', 'Foglalkoztató neve: Szautner László', header_format)
        sheet[i].write('M1', 'Munkavégzés helye: Nóráp', header_format)
        sheet[i].merge_range('A3:B5', year, year_format)
        sheet[i].write('A6', "Dátum", workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":2, "right":1, "bottom":2}))
        sheet[i].write('B6', "Napok", workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":1, "right":2, "bottom":2}))

        """Fill the date and days (A7:B68)"""
        weekend_rows = []
        for j in range(0, 2 * calendar.monthrange(year, i + 1)[1], 2):
            sheet[i].merge_range(j + 6, 0, j + 7, 0, datetime.date(year, i + 1, j // 2 + 1), date_format)
            sheet[i].merge_range(j + 6, 1, j + 7, 1, datetime.date(year, i + 1, j // 2 + 1), day_format)
        sheet[i].merge_range(j + 8, 0, j + 8, 1, "Dolgozó aláírása", format_whole)
        sheet[i].merge_range(j + 10, 0, j + 10, 1, "Összesen", format)

        """Fill the legends"""
        cell_format = workbook.add_format({"bg_color":"#0070C0"})
        sheet[i].write(j + 12, 2, "", cell_format)
        sheet[i].write(j + 12, 3, "Munkaszüneti nap", legend_format)
        cell_format = workbook.add_format({"bg_color":"#00B050"})
        sheet[i].write(j + 12, 6, "", cell_format)
        sheet[i].write(j + 12, 7, "Áthelyezett pihenőnap", legend_format)
        cell_format = workbook.add_format({"bg_color":"#FF0000"})
        sheet[i].write(j + 12, 11, "", cell_format)
        sheet[i].write(j + 12, 12, "Áthelyezett munkanap", legend_format)

        """Fill the columns of the employees"""
        col1 = -1
        col2 = 1
        for element in employees:
            """Fill the header"""
            holiday_list = []
            col1 = col1 + 3
            col2 = col2 + 3
            sheet[i].set_row(2, 10.5)
            sheet[i].set_row(3, 10.5)
            sheet[i].merge_range(2, col1, 3, col2, element.name, format_name)
            sheet[i].merge_range(4, col1, 4, col2 - 1, "Munkaidő", format_left)
            sheet[i].write(4, col2, "M. idő", format_right)
            sheet[i].write(5, col1, "Kezd.", workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":2, "right":1, "bottom":2}))
            sheet[i].write(5, col2 - 1, "Vége", workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":1, "right":1, "bottom":2}))
            sheet[i].write(5, col2, "óra", workbook.add_format({"font_name": "Segoe UI", "font_size": 8, "align": "center", "valign": "vcenter", "border_color":"#C0C0C0", "top":1, "left":1, "right":2, "bottom":2}))
            sheet[i].merge_range(j + 8, col1, j + 8, col2, "", format_whole)

            """Fill the sum of workhours"""
            sheet[i].write(j + 10, col2, "=SUM({}:{})".format(xl_rowcol_to_cell(6, col2), xl_rowcol_to_cell(j + 7, col2)), format_whole) # Summarized working hours

            start1 = time.strptime(element.workstart1, "%H:%M")
            end1 = time.strptime(element.workend1, "%H:%M")
            workhours1 = end1.tm_hour - start1.tm_hour
            for index in element.holidays:
                holiday_list.append(days_hu[str.lower(index)])

            if (element.workstart2 != "") & (element.workend2 != ""):
                start2 = time.strptime(element.workstart2, "%H:%M")
                end2 = time.strptime(element.workend2, "%H:%M")
                workhours2 = end2.tm_hour - start2.tm_hour
                for row in range(6, j + 8, 2):
                    day_cursor = datetime.date(year, i + 1, (row - 6) // 2 + 1).weekday()
                    if day_cursor in holiday_list:
                        sheet[i].merge_range(row, col1, row + 1, col1 + 2, "SZ", holiday_format)
                    else:
                        sheet[i].write(row, col1, element.workstart1, time_format_left)
                        sheet[i].write(row + 1, col1, element.workstart2, time_format_left)
                        sheet[i].write(row, col1 + 1, element.workend1, time_format)
                        sheet[i].write(row + 1, col1 + 1, element.workend2, time_format)
                        sheet[i].write(row, col1 + 2, workhours1, worktime_format)
                        sheet[i].write(row + 1, col1 + 2, workhours2, worktime_format)
            elif (element.workstart2 == "") & (element.workend2 == ""):
                for row in range(6, j + 8, 2):
                    day_cursor = datetime.date(year, i + 1, (row - 6) // 2 + 1).weekday()
                    if day_cursor in holiday_list:
                        sheet[i].merge_range(row, col1, row + 1, col1 + 2, "SZ", holiday_format)
                    else:
                        sheet[i].merge_range(row, col1, row + 1, col1, element.workstart1, time_format_left)
                        sheet[i].merge_range(row, col1 + 1, row + 1, col1 + 1, element.workend1, time_format)
                        sheet[i].merge_range(row, col1 + 2, row + 1, col1 + 2, workhours1, worktime_format)
            else:
                print("error")

        sheet[i].merge_range(j + 9, 2, j + 9, col2, "", format_whole)
    workbook.close()
    messagebox.showinfo("", "Munkaidőnyilvántartás sikeresen legenerálva!")

def input_window():
    entries = [] # List for clearing the entry fields

    """Create the input window"""
    app = tk.Tk()
    app.geometry("450x350")
    app.title("Munkaidő nyilvántartás generáló")

    """Input year"""
    tk.Label(app,text="Évszám:").place(x=40,y=30)
    year = tk.StringVar(app)
    tk.Entry(app,textvariable=year,width=10).place(x=130,y=30)

    """Input employee's name"""
    tk.Label(app, text="Dolgozó neve:").place(x=40, y=60)
    name = tk.StringVar(app)
    entries.append(tk.Entry(app, textvariable=name, width=40))
    entries[0].place(x=130, y=60)

    """Input start and end of the work"""
    tk.Label(app, text="Munka kezdés:").place(x=40, y=90)
    workstart1 = tk.StringVar(app)
    entries.append(tk.Entry(app, textvariable=workstart1, width=10))
    entries[1].place(x=130, y=90)
    tk.Label(app, text="Munka vége:").place(x=230, y=90)
    workend1 = tk.StringVar(app)
    entries.append(tk.Entry(app, textvariable=workend1, width=10))
    entries[2].place(x=310, y=90)

    tk.Label(app, text="Munka kezdés:").place(x=40, y=120)
    workstart2 = tk.StringVar(app)
    entries.append(tk.Entry(app, textvariable=workstart2, width=10))
    entries[3].place(x=130, y=120)
    tk.Label(app, text="Munka vége:").place(x=230, y=120)
    workend2 = tk.StringVar(app)
    entries.append(tk.Entry(app, textvariable=workend2, width=10))
    entries[4].place(x=310, y=120)

    """Input holidays"""
    tk.Label(app, text="Szabadnapok:").place(x=40, y=150)
    days = ('Hétfő', 'Kedd', 'Szerda', 'Csütörtök', 'Péntek', 'Szombat', 'Vasárnap')
    holiday_items = tk.Variable(value=days)
    listbox = tk.Listbox(app, listvariable=holiday_items, height=7, selectmode=tk.MULTIPLE)
    listbox.place(x=130, y=150)

    """Collect the content of the inputs"""
    content = [name, workstart1, workend1, workstart2, workend2]

    """Action buttons: add new employee button, creat the excel button and exit button"""
    add_button=tk.Button(app,text="Új dolgozó",command=lambda: add_employee(entries, content, listbox)).place(x=80, y=285)
    export_button = tk.Button(app, text="Generálás", command=lambda: create_excel(employees, int(year.get()))).place(x=200, y=285)
    exit_button = tk.Button(app, text="Kilépés", command="exit").place(x=300, y=285)

    app.mainloop()

def add_employee(entry_list, emp_data, listbox):
    holidays = [] # List of the selected holidays

    for i in listbox.curselection():
        holidays.append(listbox.get(i))

    """Fill the employees list with the data of all employees"""
    employees.append(Employee(emp_data[0].get(), emp_data[1].get(), emp_data[2].get(), emp_data[3].get(), emp_data[4].get(), holidays))
    messagebox.showinfo("", "{} sikeresen hozzáadva!".format(emp_data[0].get()))
    for elements in entry_list:
        elements.delete(0,'end')
    listbox.selection_clear(0,'end')

employees = [] # Global list for items of Employee class

"""FINAL"""
input_window() # Call of GUI

"""TEST"""
#employees = [Employee("Varga János", "7:00", "15:00", "17:00", "20:00", ["Kedd", "Vasárnap", "Péntek"]), Employee("Varga András", "10:00", "12:00", "13:00", "16:00", ["Szerda", "Péntek"]), Employee("Süle Imre", "8:00", "16:00", "", "", ["Szombat", "Vasárnap", "Hétfő", "Kedd", "Szerda", "Csütörtök"])]
#create_excel(employees, 2023)
