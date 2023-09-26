"""
GUI of degree days/hours calculator
Requires Internet connection
Weather data source: meteostat.net
"""

import tkinter as tk
import sys
sys.path.append('..')
from Shared.IAC import degree_days, degree_hours

def calculate():
    try:
        # get the values from the GUI
        zipcode = entry_zip.get()
        mode = Mode.get().lower()
        calctype = CalcType.get().lower().split()[1]
        basetemp = entry_basetemp.get()
        setback = entry_setback.get()
        history = int(drop_clicked.get().split()[0])
        if basetemp.isdigit():
            basetemp = int(entry_basetemp.get())
        else:
            raise Exception("Base temperature must be a valid integer")
        if setback.isdigit():
            setback = int(entry_setback.get())
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
        # convert schedule to a tuple of tuples
        schedule = tuple(schedule)
        # call functions from IAC.py
        if calctype == "hours":
            result = degree_hours(zipcode, mode, basetemp, setback, schedule, history)
        else:
            result = degree_days(zipcode, mode, basetemp, history)
        text_result.config(state='normal')
        text_result.delete('1.0', tk.END)
        # format result to integer with thousands separator
        text_result.insert(tk.END, "{:,}".format(int(result)))
        text_result.config(state='disabled')
    except:
        # show a pop-up window if there is an error
        popup = tk.Tk()
        popup.wm_title("Error")
        # The error message is the exception message
        label = tk.Label(popup, text=sys.exc_info()[1], font=('Arial', 14))
        label.pack(side="top", fill="x", pady=10)
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
                frame_schedule.grid_slaves(row=j+1,column=i+1)[0].config(state='disabled')
    else:
        # enable setback temp
        entry_setback.config(state='normal')
        # restore schedule grid widgets
        check_hours()

def check_hours():
    for i in range(7):
        dropdown_start = frame_schedule.grid_slaves(row=1,column=i+1)[0]
        dropdown_end = frame_schedule.grid_slaves(row=2,column=i+1)[0]
        check_allday = frame_schedule.grid_slaves(row=3,column=i+1)[0]
        check_holiday = frame_schedule.grid_slaves(row=4,column=i+1)[0]
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
outpad = 10
inpad = 5
# GUI title
label_iac = tk.Label(window, text="IAC Degree Days/Hours Calculator", font=('Arial', 24))
label_iac.pack()
# Left frame
frame_left = tk.Frame(window, highlightbackground="gray", highlightthickness=1)
# ZIP frame
frame_zip = tk.Frame(frame_left, width=20)
# ZIP code lable
label_zip = tk.Label(frame_zip, text="ZIP Code", width=10, anchor='w')
label_zip.pack(side='left')
# ZIP code entry
entry_zip = tk.Entry(frame_zip,width=10)
entry_zip.insert(0, "18015")
entry_zip.pack(side='right')
frame_zip.pack(padx=inpad, pady=inpad)

# Mode frame
frame_mode = tk.Frame(frame_left)
# Mode lable
label_mode = tk.Label(frame_mode, text="Mode", width=10, anchor='w')
label_mode.pack(side='left')
# Radio button frame
frame_radio1 = tk.Frame(frame_mode)
# Mode radio buttons
Mode = tk.StringVar()
Mode.set("Cooling")
radiocool = tk.Radiobutton(frame_radio1, text="Cooling", variable=Mode, value="Cooling", command=updatewidget, width=10)
radiocool.pack(side='top')
radioheat = tk.Radiobutton(frame_radio1, text="Heating", variable=Mode, value="Heating", command=updatewidget, width=10)
radioheat.pack(side='bottom')
frame_radio1.pack(side='right')
frame_mode.pack(padx=inpad, pady=inpad)

# Base temp frame
frame_basetemp = tk.Frame(frame_left)
# Base temp lable
label_basetemp = tk.Label(frame_basetemp, text="Base Temp.", width=10, anchor='w')
label_basetemp.pack(side='left')
# Base temp entry
entry_basetemp = tk.Entry(frame_basetemp,width=10)
entry_basetemp.insert(0, "65")
entry_basetemp.pack(side='right')
frame_basetemp.pack(padx=inpad, pady=inpad)

# Setback temp frame
frame_setback = tk.Frame(frame_left)
# Setback temp lable
label_setback = tk.Label(frame_setback, text="Setback Temp.", width=10, anchor='w')
label_setback.pack(side='left')
# Setback temp entry
entry_setback = tk.Entry(frame_setback,width=10)
entry_setback.insert(0, "65")
entry_setback.pack(side='right')
frame_setback.pack(padx=inpad, pady=inpad)

# History frame
frame_history = tk.Frame(frame_left)
# History lable
label_history = tk.Label(frame_history, text="History", width=10, anchor='w')
label_history.pack(side='left')
# History dropdown menu
drop_options = ["1 year","2 years","3 years","4 years","5 years"]
drop_clicked = tk.StringVar()
drop_clicked.set("4 years")
drop_history = tk.OptionMenu(frame_history, drop_clicked, *drop_options)
drop_history.config(width=6, anchor='w')
drop_history.pack(side='right')
frame_history.pack(padx=inpad, pady=inpad)

# Calculation type frame
frame_type = tk.Frame(frame_left)
# Calculation type lable
label_dayhour = tk.Label(frame_type, text="Calc. Type", width=10, anchor='w')
label_dayhour.pack(side='left')
# Radiobutton frame
frame_radio2 = tk.Frame(frame_type)
# Calculation type radio buttons
CalcType = tk.StringVar()
CalcType.set("Degree Hours")
radioday = tk.Radiobutton(frame_radio2, text="Deg. Days", variable=CalcType, value="Degree Days", command=updatewidget, width=10, anchor='w')
radioday.pack(side='top')
hourmode = tk.StringVar()
radiohour = tk.Radiobutton(frame_radio2, text="Deg. Hours", variable=CalcType, value="Degree Hours", command=updatewidget, width=10, anchor='w')
radiohour.pack(side='bottom')
frame_radio2.pack(side='right')
frame_type.pack(padx=inpad, pady=inpad)

frame_left.pack(side='left',padx=outpad, pady=outpad)

# Right frame
frame_right = tk.Frame(window, width=20, highlightbackground="gray", highlightthickness=1)

# Schedule frame
label_schedule = tk.Label(frame_right, text="Thermostat Programming Schedule", font=('Arial', 20))
label_schedule.pack()
frame_schedule = tk.Frame(frame_right)
tk.Label(frame_schedule, text="", width=5).grid(row=0, column=0, pady=inpad)
tk.Label(frame_schedule, text="Start", width=5, anchor='w').grid(row=1, column=0, pady=inpad)
tk.Label(frame_schedule, text="End", width=5, anchor='w').grid(row=2, column=0, pady=inpad)
tk.Label(frame_schedule, text="All Day", width=5, anchor='w').grid(row=3, column=0, pady=inpad)
tk.Label(frame_schedule, text="Holiday", width=5, anchor='w').grid(row=4, column=0, pady=inpad)
days=['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
hours=list(range(0,25))
start_list = []
end_list = []
allday_list = []
holiday_list = []
for i in range(7):
    # Day label
    tk.Label(frame_schedule, text=days[i], width=5, anchor='w').grid(row=0, column=i+1, pady=inpad)
    # Start time entry
    start_var = tk.StringVar()
    start_var.set("9")
    start_list.append(start_var)
    dropdown_start = tk.OptionMenu(frame_schedule, start_var, *hours)
    dropdown_start.config(width=1)
    dropdown_start.grid(row=1, column=i+1)
    # End time entry
    end_var = tk.StringVar()
    end_var.set("17")
    end_list.append(end_var)
    dropdown_end = tk.OptionMenu(frame_schedule, end_var, *hours)
    dropdown_end.config(width=1)
    dropdown_end.grid(row=2, column=i+1)
    # 24 hr check box
    allday_var = tk.IntVar()
    allday_list.append(allday_var)
    checkbox_allday = tk.Checkbutton(frame_schedule, variable=allday_var, width=2, command=check_hours)
    checkbox_allday.grid(row=3, column=i+1)
    # holiday check box
    holiday_var = tk.IntVar()
    holiday_list.append(holiday_var)
    checkbox_holiday = tk.Checkbutton(frame_schedule, variable=holiday_var, width=2, command=check_hours)
    checkbox_holiday.grid(row=4, column=i+1)
frame_schedule.pack(side='top',padx=inpad, pady=inpad)

frame_result = tk.Frame(frame_right)
# Calculate botton
button_calc = tk.Button(frame_result, text ="Calculate", width=6, command = calculate, font=('Arial', 18))
button_calc.pack(side='left')
# Result textbox
text_result = tk.Text(frame_result, state='disabled', width=8, height=1, font=('Arial', 18))
text_result.pack(side='right')
# Result label
resultlabel= tk.StringVar()
resultlabel.set("Cooling Degree Hours:")
label_result = tk.Label(frame_result, textvariable = resultlabel, width=18, font=('Arial', 18))
label_result.pack(side='right')
frame_result.pack(padx=inpad, pady=inpad)

label_copyright = tk.Label(frame_right, text="© 2023 Lehigh University Industrial Assessment Center")
label_copyright.pack()
frame_right.pack(side='right', padx=outpad, pady=outpad)
window.mainloop()