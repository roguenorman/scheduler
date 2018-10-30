import win32com.client
import datetime as dt
import time
import pywintypes
import win32timezone
import configparser
import gui
import os
 
def get_config():
    config = configparser.ConfigParser()
    home = os.path.expanduser("~")
    path = os.path.join(home, 'Documents', 'scheduler.conf')
    config.read(path)
    conf = config['DEFAULT']
    start = conf.get('start', fallback='8:00')
    end = conf.get('end', fallback='17:00')
    duration = int(conf.get('duration', fallback='1'))
    period = int(conf.get('period', fallback='5'))
    days = conf.get('days', fallback='(0, 1, 2, 3, 4)')
    days = tuple([int(i) for i in days[1:-1].split(',')])
    return (start, end, duration, period, days)

#get appointments from a period
def get_appts(calendar, day_start, day_end, appt_period, work_days):
    calendar.Items.Sort("[Start]")
    calendar.Items.IncludeRecurrences = True   
    appt_list = []
    today = dt.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    date_list = [today + dt.timedelta(days=x) for x in range(0, appt_period)]
    # dstart_splt = day_start.replace(':', " ").split()
    dstart_splt = day_start.split(":")
    # dend_splt = day_end.replace(':', " ").split()
    dend_splt = day_end.split(":")

    for date in date_list:
        if date.weekday() in work_days:
            #add day start
            #work_hours = (date.replace(hour=int(dstart_splt[0]), minute=int(dstart_splt[1]), second=0, microsecond=00), date.replace(hour=int(dend_splt[0]), minute=int(dend_splt[1]), second=00, microsecond=00))
            work_hours = (date.replace(hour=int(dstart_splt[0]), minute=int(dstart_splt[1])), date.replace(hour=int(dend_splt[0]), minute=int(dend_splt[1])))
            appt_list.append((work_hours[0],work_hours[0]))
            #add appointments
            filter = "[Start] >= '" + date.strftime("%d %m %Y") + " " + day_start + "' AND [Start] <= '" + date.strftime("%d %m %Y") + " " + day_end + "'"
            # filter = "[Start] >= '" + date.strftime("%d %m %Y") + " " + day_start + "' AND [Start] <= '" + date.strftime("%d %m %Y") + " " + day_end + "'"
            results = calendar.Items.Restrict(filter)
            for appt in results:
                if appt.IsRecurring:
                    rp = appt.GetRecurrencePattern()
                    # get first occurance
                    start = appt.Start.strptime(appt.Start.Format() + ' +00:00', '%a %b %d %H:%M:%S %Y %z')
                    # replace first occurance with current occurance
                    appt_start = start.replace(day=date.day, month=date.month, year=date.year)
                    appt.Close(1)
                    appt = None
                    appt = rp.GetOccurrence(appt_start)

                appt_start = dt.datetime.strptime(appt.Start.Format(), '%a %b %d %H:%M:%S %Y')
                appt_end = dt.datetime.strptime(appt.End.Format(), '%a %b %d %H:%M:%S %Y')
                # restrict is returning items outside of the filtered time so we need to remove them
                if work_hours[0] <= appt_start < work_hours[1]:
                    appt_list.append((appt_start, appt_end))
            #add day end
            appt_list.append((work_hours[1],work_hours[1]))

    return appt_list
 
def get_slots(appointments, appt_duration):
    duration = dt.timedelta(hours=appt_duration)
    free_list = []
    slots = sorted(appointments)
    for start, end in ((slots[i][1], slots[i+1][0]) for i in range(len(slots)-1)):
        #assert start <= end, "Cannot attend all appointments"
        if start.weekday() == end.weekday():
            if start + duration <= end:
                free_list.append("{:%A %d %B %Y} - from {:%I:%M%p} until {:%I:%M%p}".format(start, start, end))
                start += duration
    return free_list

def create_email(outlook, body):
    mail = outlook.CreateItem(0)
    mail.Body = str(body)
    mail.Display(1)


def get_availability():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)

    get_config()
    day_start, day_end, appt_duration, appt_period, work_days = get_config()
    appointments = get_appts(calendar, day_start, day_end, appt_period, work_days)
    #print(sorted(appointments))
    slots = get_slots(appointments, appt_duration)
    body = '\n'.join(slots)

    create_email(outlook, body)
 
get_availability()