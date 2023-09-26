"""
GUI of degree days/hours calculator
Optimized for macOS
Requires Internet connection
Weather data source: meteostat.net
"""

import tkinter as tk
import sys, pgeocode, webbrowser
sys.path.append('..')
from Shared.IAC import degree_days, degree_hours

def calculate():
    try:
        # get the values from the GUI
        zipcode = entry_zip.get()
        location = pgeocode.Nominatim('us').query_postal_code(zipcode)
        Address.set(location.place_name + ', ' + location.state_code)
        mode = Mode.get().lower()
        calctype = CalcType.get().lower().split()[1]
        basetemp = entry_basetemp.get()
        setback = entry_setback.get()
        history = int(drop_clicked.get().split()[0])
        if basetemp.isdigit():
            basetemp = int(basetemp)
        else:
            raise Exception("Base temperature must be a valid integer")
        if setback.isdigit():
            setback = int(setback)
        else:
            raise Exception("Setback temperature must be a valid interger")
        # schedule is a tuple of 7 tuples, each tuple is a pair of start and end time
        schedule = []
        for i in range(7):
            start = int(start_list[i].get())
            end = int(end_list[i].get())
            if allday_list[i].get() == 1:
                schedule.append((0,24))
            elif holiday_list[i].get() == 1:
                schedule.append((0,0))
            else:
                schedule.append((start,end))
        # convert from list to tuple
        schedule = tuple(schedule)
        # call functions from IAC.py
        if calctype == "hours":
            result = degree_hours(zipcode, mode, basetemp, setback, schedule, history)
        else:
            result = degree_days(zipcode, mode, basetemp, history)
        text_result.config(state='normal')
        text_result.delete('1.0', tk.END)
        # format result to integer with thousand separator
        text_result.insert(tk.END, "{:,}".format(int(result)))
        text_result.config(state='disabled')
    except:
        # show a pop-up window if there is an error
        popup = tk.Tk()
        # center the pop-up window
        popup.eval('tk::PlaceWindow . center')
        popup.wm_title("Error")
        # The error message is the exception message
        label = tk.Label(popup, text=sys.exc_info()[1])
        label.pack(side="top", fill="x", padx=pad)
        # close the pop-up window
        button = tk.Button(popup, text="OK", command = popup.destroy)
        button.pack()
        popup.mainloop()
    
def updatewidget():
    resultlabel.set(Mode.get()+ " " + CalcType.get() + ":")
    if CalcType.get() == "Degree Days":
        # disable setback temp
        entry_setback.config(state='disabled')
        # disable all schedule grid widgets
        for i in range(7):
            for j in range(4):
                frame_right.grid_slaves(row=j+2,column=i+1)[0].config(state='disabled')
    else:
        # enable setback temp
        entry_setback.config(state='normal')
        # restore schedule grid widgets
        check_hours()

def check_hours():
    for i in range(7):
        dropdown_start = frame_right.grid_slaves(row=2,column=i+1)[0]
        dropdown_end = frame_right.grid_slaves(row=3,column=i+1)[0]
        check_allday = frame_right.grid_slaves(row=4,column=i+1)[0]
        check_holiday = frame_right.grid_slaves(row=5,column=i+1)[0]
        # if the check_allday box is checked
        if allday_list[i].get() == 1:
            # disable other dropdowns
            dropdown_start.config(state='disabled')
            dropdown_end.config(state='disabled')
            check_allday.config(state='normal')
            check_holiday.config(state='disabled')
        # if the check_holiday box is checked
        elif holiday_list[i].get() == 1:
            # disable other dropdowns
            dropdown_start.config(state='disabled')
            dropdown_end.config(state='disabled')
            check_allday.config(state='disabled')
            check_holiday.config(state='normal')
        else:
            # enable everything
            dropdown_start.config(state='normal')
            dropdown_end.config(state='normal')
            check_allday.config(state='normal')
            check_holiday.config(state='normal')

# initialize GUI
window = tk.Tk()
# Center the window
window.eval('tk::PlaceWindow . center')
# GUI title
window.title("IAC Degree Days/Hours Calculator")
pad = 5

# Left frame
frame_left = tk.Frame(window)
leftw1=10
leftw2=10
Address = tk.StringVar()
Address.set("Bethlehem, PA")
tk.Label(frame_left, text="ZIP Code", width=leftw1, anchor='w').grid(row=0, column=0, pady=pad)
tk.Label(frame_left, textvariable=Address, width=leftw1+leftw2, anchor='e').grid(row=1, column=0, columnspan=2, pady=pad)
tk.Label(frame_left, text="Mode", width=leftw1, anchor='w').grid(rowspan=2, column=0, pady=pad)
tk.Label(frame_left, text="Base Temp.", width=leftw1, anchor='w').grid(row=4, column=0, pady=pad)
tk.Label(frame_left, text="Setback Temp.", width=leftw1, anchor='w').grid(row=5, column=0, pady=pad)
tk.Label(frame_left, text="History", width=leftw1, anchor='w').grid(row=6, column=0, pady=pad)
tk.Label(frame_left, text="Method", width=leftw1, anchor='w').grid(rowspan=2, column=0, pady=pad)


entry_zip = tk.Entry(frame_left, width=leftw2)
entry_zip.insert(0, "18015")
entry_zip.grid(row=0, column=1, pady=pad, sticky='w')

Mode = tk.StringVar()
Mode.set("Cooling")
radiocool = tk.Radiobutton(frame_left, text="Cooling", variable=Mode, value="Cooling", command=updatewidget, width=leftw2, anchor='w')
radiocool.grid(row=2, column=1, sticky='w')
radioheat = tk.Radiobutton(frame_left, text="Heating", variable=Mode, value="Heating", command=updatewidget, width=leftw2, anchor='w')
radioheat.grid(row=3, column=1, sticky='w')

entry_basetemp = tk.Entry(frame_left, width=leftw2)
entry_basetemp.insert(0, "65")
entry_basetemp.grid(row=4, column=1, pady=pad, sticky='w')

entry_setback = tk.Entry(frame_left, width=leftw2)
entry_setback.insert(0, "65")
entry_setback.grid(row=5, column=1, pady=pad, sticky='w')

drop_options = ["1 year","2 years","3 years","4 years","5 years"]
drop_clicked = tk.StringVar()
drop_clicked.set("4 years")
drop_history = tk.OptionMenu(frame_left, drop_clicked, *drop_options)
drop_history.config(width=leftw2-4, anchor='w')
drop_history.grid(row=6, column=1, pady=pad, sticky='w')

CalcType = tk.StringVar()
CalcType.set("Degree Hours")
radioday = tk.Radiobutton(frame_left, text="Deg. Days", variable=CalcType, value="Degree Days", command=updatewidget, width=leftw2, anchor='w')
radioday.grid(row=7, column=1, sticky='w')
radiohour = tk.Radiobutton(frame_left, text="Deg. Hours", variable=CalcType, value="Degree Hours", command=updatewidget, width=leftw2, anchor='w')
radiohour.grid(row=8, column=1, sticky='w')

frame_left.pack(side='left', padx=2*pad, pady=pad)

# Right frame
frame_right = tk.Frame(window)

rightw1=6
tk.Label(frame_right, text="Thermostat Programming Schedule").grid(row=0, columnspan=8, pady=pad)
tk.Label(frame_right, text="Start", width=rightw1, anchor='w').grid(row=2, column=0, pady=pad)
tk.Label(frame_right, text="End", width=rightw1, anchor='w').grid(row=3, column=0, pady=pad)
tk.Label(frame_right, text="All Day", width=rightw1, anchor='w').grid(row=4, column=0, pady=pad)
tk.Label(frame_right, text="Holiday", width=rightw1, anchor='w').grid(row=5, column=0, pady=pad)

days=['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
hours=list(range(0,25))
start_list = []
end_list = []
allday_list = []
holiday_list = []
for i in range(7):
    # Day label
    tk.Label(frame_right, text=days[i]).grid(row=1, column=i+1)
    # Start time entry
    start_var = tk.StringVar()
    start_var.set("9")
    start_list.append(start_var)
    dropdown_start = tk.OptionMenu(frame_right, start_var, *hours)
    dropdown_start.config(width=1)
    dropdown_start.grid(row=2, column=i+1)
    # End time entry
    end_var = tk.StringVar()
    end_var.set("17")
    end_list.append(end_var)
    dropdown_end = tk.OptionMenu(frame_right, end_var, *hours)
    dropdown_end.config(width=1)
    dropdown_end.grid(row=3, column=i+1)
    # 24 hr check box
    allday_var = tk.IntVar()
    allday_list.append(allday_var)
    checkbox_allday = tk.Checkbutton(frame_right, variable=allday_var, command=check_hours)
    checkbox_allday.grid(row=4, column=i+1)
    # holiday check box
    holiday_var = tk.IntVar()
    holiday_list.append(holiday_var)
    checkbox_holiday = tk.Checkbutton(frame_right, variable=holiday_var, command=check_hours)
    checkbox_holiday.grid(row=5, column=i+1)

# Calculate botton
button_calc = tk.Button(frame_right, text ="Calculate", width=6, command = calculate)
button_calc.grid(row=6, column=1, columnspan=2, pady=pad)

# Result label
resultlabel= tk.StringVar()
resultlabel.set("Cooling Degree Hours:")
label_result = tk.Label(frame_right, textvariable = resultlabel)
label_result.grid(row=6, column=3, columnspan=3, pady=pad, sticky='w')

# Result textbox
text_result = tk.Text(frame_right, state='disabled', height=1, width=10)
text_result.grid(row=6, column=6, columnspan=2, pady=pad, sticky="we")

# Data source label
label_datasource = tk.Label(frame_right, text="Weather Data Source: Meteostat.net", fg="blue", cursor="hand2")
# hyperlink to meteostat.net/en
label_datasource.bind("<Button-1>", lambda e: webbrowser.open_new_tab("https://meteostat.net/en"))
label_datasource.grid(row=7, columnspan=8, pady=pad)

# Copyright label
label_copyright = tk.Label(frame_right, text="© 2023 Lehigh University Industrial Assessment Center", fg="blue", cursor="hand2")
# hyperlink to iac.lehigh.edu
label_copyright.bind("<Button-1>", lambda e: webbrowser.open_new_tab("https://luiac.cc.lehigh.edu"))
label_copyright.grid(row=8, columnspan=8)

frame_right.pack(side='right', padx=2*pad, pady=pad)
window.mainloop()