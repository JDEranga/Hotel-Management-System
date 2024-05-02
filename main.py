import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import openpyxl


def load_data():
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)


def insert_row():
    name = name_entry.get()
    nic = nic_entry.get()
    num = num_entry.get()
    count = int(count_spinbox.get())
    room_status = status_combobox.get()
    date = cal.get_date()

    print(name, nic, num, count, room_status, date)

    # Insert row into Excel sheet
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, nic, num, count, room_status, date]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    nic_entry.delete(0, "end")
    nic_entry.insert(0, "NIC")
    num_entry.delete(0, "end")
    num_entry.insert(0, "Phone")
    count_spinbox.delete(0, "end")
    count_spinbox.insert(0, "No. of Guests")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])
    cal.set_date(None)




def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

combo_list = ["Double Room", "Triple Room", "Family Room"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Details")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

#Insert

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Guest Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

nic_entry = ttk.Entry(widgets_frame)
nic_entry.insert(0, "NIC")
nic_entry.bind("<FocusIn>", lambda e:nic_entry.delete('0', 'end'))
nic_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

num_entry = ttk.Entry(widgets_frame)
num_entry.insert(0, "Phone")
num_entry.bind("<FocusIn>", lambda e:num_entry.delete('0', 'end'))
num_entry.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

count_spinbox = ttk.Spinbox(widgets_frame, from_=0, to=100)
count_spinbox.insert(0, "No. of Guests")
count_spinbox.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=4, column=0, padx=5, pady=5,  sticky="ew")


cal = DateEntry(widgets_frame, width=12, background='darkblue', foreground='white', borderwidth=2)
cal.grid(row=5, column=0, padx=5, pady=5, sticky="ew")






button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=7, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=8, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(
    widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=9, column=0, padx=5, pady=10, sticky="nsew")


treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Name", "NIC", "Phone", "Count", "Room","Date")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Name", width=100)
treeview.column("NIC", width=100)
treeview.column("Phone", width=100)
treeview.column("Count", width=50)
treeview.column("Room", width=100)
treeview.column("Date", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()


root.mainloop()
