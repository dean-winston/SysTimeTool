import sys
import win32api
import ctypes, sys
import time
import win32api
import socket
import struct
import csv

from threading import Timer
from tkinter import *
from tkinter import ttk
from datetime import datetime, timedelta


server_list = ['ntp.iitb.ac.in', 'time.nist.gov', 'time.windows.com', 'pool.ntp.org']
MAX_COLOUM = 7
UTC_ADD = 8

global g_index
global g_timer
global time_list

def update_set():
    global g_index
    global g_timer
    global time_list
    global UTC_ADD
    print("update_set", g_index, len(time_list))
    if g_index >= len(time_list):
        return

    curSetting = time_list[g_index]
    dateSetting = curSetting[1]
    if dateSetting:
        dt = datetime(*dateSetting)
        dt = dt + timedelta(hours=UTC_ADD)
        win32api.SetLocalTime(dt)
    
    delaySec = curSetting[0]
    if delaySec > 0:
        g_index = g_index + 1
        g_timer = Timer(delaySec, update_set)
        g_timer.start()
    else:
        close_set()

def start_set():
    print("start")
    reload_setting()
    global g_index
    g_index = 0
    update_set()

def close_set():
    global g_timer
    print("closeing")
    if g_timer and g_timer.is_alive():
        g_timer.cancel()
    g_timer = None

def restart_set():
    close_set()
    start_set()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_admin():
    if is_admin():
        main_window()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)

def gettime_ntp(addr='time.nist.gov'):
    TIME1970 = 2208988800
    client = socket.socket( socket.AF_INET, socket.SOCK_DGRAM )
    data = '\x1b' + 47 * '\0'
    data = data.encode()
    try:
        client.settimeout(5.0)
        client.sendto( data, (addr, 123))
        data, address = client.recvfrom( 1024 )
        if data:
            epoch_time = struct.unpack( '!12I', data )[10]
            epoch_time -= TIME1970
            return epoch_time
    except socket.timeout:
        return None

def sync_world_time():
    for server in server_list:
        epoch_time = gettime_ntp(server)
        if epoch_time is not None:
            utcTime = datetime.utcfromtimestamp(epoch_time)
            win32api.SetSystemTime(utcTime.year, utcTime.month, utcTime.weekday(), utcTime.day, utcTime.hour, utcTime.minute, utcTime.second, 0)
            localTime = datetime.fromtimestamp(epoch_time)
            print("Time updated to: " + localTime.strftime("%Y-%m-%d %H:%M") + " from " + server)
            break
        else:
            print("Could not find time from " + server)

def reload_setting():
    global time_list
    time_list = []
    with open('setting.csv', encoding = 'UTF-8') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
        i = 0
        for row in spamreader:
            print("line", i, row)
            i = i + 1
            if i > 1:
                numRow = []
                for param in row:
                    numRow.append(int(param))

                print(numRow[3])
                numRow[3] = numRow[3]
                print(numRow[3])
                
                dateList = []
                for x in range(MAX_COLOUM):
                    dateList.append(numRow[x])

                dataLen = len(numRow)
                curlist = []
                if dataLen >= MAX_COLOUM:
                    delaySec = numRow[MAX_COLOUM]
                    curlist = [delaySec, dateList]
                else:
                    curlist = [0, dateList]
                time_list.append(curlist)
                
                print(time_list)

def main_window():
    root = Tk()
    frm = ttk.Frame(root, padding=100)
    frm.grid()
    
    ttk.Button(frm, text="Start", command=start_set).grid(column=1, row=0)
    ttk.Button(frm, text="Stop", command=close_set).grid(column=2, row=0)
    ttk.Button(frm, text="Restart", command=restart_set).grid(column=3, row=0)
    ttk.Button(frm, text="Recover", command=sync_world_time).grid(column=4, row=0)

    root.title("SysTimeTool")
    root.mainloop()

if __name__ == "__main__":
    check_admin()