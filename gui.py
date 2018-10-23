import tkinter as tk
import datetime as dt
import configparser
import os
from tkinter import ttk


config = configparser.ConfigParser()
days = [(0, 'Monday'), (1, 'Tuesday'), (2, 'Wednesday'), (3, 'Thursday'), (4, 'Friday'), (5, 'Saturday'), (6, 'Sunday')]
# hours = [(i, dt.time(i).strftime("%H:%M"))for i in range(24)]

time = dt.datetime.strptime('00:00','%H:%M')
hours = [(i, (time + dt.timedelta(minutes=30*i)).strftime('%H:%M')) for i in range(0, 48)]


def save_config(start, end, duration, period, selected_days):
    #saves the settings file
    config.set('DEFAULT', 'Start', start)
    config.set('DEFAULT', 'End', end)
    config.set('DEFAULT', 'Duration', duration)
    config.set('DEFAULT', 'Period', period)
    config.set('DEFAULT', 'Days', selected_days)
    with open('config.ini', 'w') as configfile:
        config.write(configfile)
    

def get_config():
    global config
    config.read("config.ini")
    conf = config['DEFAULT']
    start = conf.get('Start', fallback='8:00')
    end = conf.get('End', fallback='17:00')
    duration = conf.get('Duration', fallback='1')
    period = conf.get('Period', fallback='5')
    selected_days = tuple([int(i) for i in conf.get('Days', fallback='(0, 1, 2, 3, 4)')[1:-1].split(',')])
    return (start, end, duration, period, selected_days)


def build_window(start, end, duration, period, selected_days):
    #Window
    window = tk.Tk()
    window.resizable(False, False)
    window.columnconfigure(1, weight=1)
    window.columnconfigure(2, weight=1)

    window.title('Scheduler settings')
    window.geometry('335x270')

    #Option menus
    var_start = tk.StringVar(name="start")
    var_start.set(start)
    menu_start = tk.ttk.Combobox(window, textvariable=var_start, values=[*[x[1] for x in hours]])
    menu_start.config(width=12)
    menu_start.grid(column=2, row=0, sticky='e', padx=4)
    
    var_end = tk.StringVar(name="end")
    var_end.set(end)
    menu_end = tk.ttk.Combobox(window, textvariable=var_end, values=[*[x[1] for x in hours]])
    menu_end.config(width=12)
    menu_end.grid(column=2, row=1, sticky='e', padx=4)

    #Entry
    entry_dur = tk.Entry(window)
    entry_dur.grid(column=2, row=3, sticky='e', padx=5)
    entry_dur.config(width=15)
    entry_dur.insert(0, duration)


    entry_period = tk.Entry(window)
    entry_period.grid(column=2, row=2, sticky='e', padx=5)
    entry_period.config(width=15)
    entry_period.insert(0, period)

    #List box
    list_days = tk.Listbox(window, selectmode=tk.MULTIPLE, height=7)
    list_days.config(width=15)
    for index, day in days:
        list_days.insert(index, day)
    list_days.grid(column=2, row=4, sticky='e', padx=5, pady=3)
    for i in selected_days:
        list_days.select_set(i)

    #Labels
    lbl_start = tk.Label(window, text="Work Start", width=18, anchor='w')
    lbl_start.grid(column=0, row=0)
    lbl_end = tk.Label(window, text="Work End", width=18, anchor='w')
    lbl_end.grid(column=0, row=1)
    lbl_duration = tk.Label(window, text="Calendar Period (days)", width=18, anchor='w')
    lbl_duration.grid(column=0, row=2, pady=3)
    lbl_duration = tk.Label(window, text="Appt. Duration (hours)", width=18, anchor='w')
    lbl_duration.grid(column=0, row=3, pady=3)
    lbl_days = tk.Label(window, text="Work Days", width=18, anchor='w')
    lbl_days.grid(column=0, row=4)

    #Buttons
    ok_button = tk.Button(window, text="OK", command=lambda: [save_config(var_start.get(), var_end.get(), str(entry_dur.get()), entry_period.get(), str(list_days.curselection())), window.destroy()]) 
    ok_button.config(width=11)
    ok_button.grid(column=1, row=5, sticky='e', padx=5, pady=5)

    close_button = tk.Button(window, text="Close", command=window.destroy)
    close_button.config(width=11)
    close_button.grid(column=2, row=5, padx=5, pady=5)

    return window

def show_window():
    start, end, duration, period, selected_days = get_config()
    window = build_window(start, end, duration, period, selected_days)
    window.mainloop()
