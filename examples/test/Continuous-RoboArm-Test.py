import os, os.path, sys
import string

from datetime import date

import serial
from serial.tools import list_ports

import multiprocessing.dummy as mp 
import threading

import gspread
import pandas as pd

from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth
from oauth2client.service_account import ServiceAccountCredentials

import time
from uarm import swift
from uarm.wrapper.swift_api import SwiftAPI

import os, os.path, sys

# Authorizing Credentials for google sheets file

    # define the scope
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    # add credentials to the service_account
creds = ServiceAccountCredentials.from_json_keyfile_name('envelopes-project-c271f0e460b4.json',scope)
print(creds)

    # authorize the clientsheet
client = gspread.authorize(creds)

gauth = GoogleAuth()
gauth.credentials = creds
drive = GoogleDrive(gauth)

sheet = client.open("Envelopes Tracker")

tracker_worksheet = sheet.get_worksheet(0)
tracker_data = tracker_worksheet.get_all_records()
p_df = pd.DataFrame.from_dict(tracker_data)

who = ""
write_on = ""
site = ""

pod_dict = {
    "pod1": {
        "robot": "COM18",
        "plotter1": "COM50",
        "plotter2": "COM51",
        "type": "Envelope",
        "full": 600,
        "half": 100,
        "numberOfPlotters": 2, 
        "numberOfStacks": 2
    },
    "pod2": {
        "robot": "COM18",
        "plotter1": "COM22",
        "plotter2": "COM26", 
        "type": "Envelope",
        "full": 300,
        "half": 175,
        "numberOfPlotters": 2
    }
}

# Creates an array/list of command values when starting the python script
commandList = sys.argv
# Output for commands
print(commandList) 

# Initialize variables from input commands
count = 0
testcommand = commandList[1]
uport = pod_dict[commandList[2]]['robot']
current_count = 300

# Connect to uarm robot and print device information
swift = SwiftAPI(port=uport)
swift.waiting_ready()
device_info = swift.get_device_info()
print(device_info)
print(swift.port) # Port Number

# Initialize variables for AxiDraw System and print path directory for svg plot files
if sys.argv[3] is not None:
    who = sys.argv[3]
    who = who.replace("-", " ")
    write_on = sys.argv[4]
    site = sys.argv[5]
    print(f"{who},{write_on},{site}")
else:
    who, write_on, site = input("Who? E or I? Site?").split(",")

plot_path = ''

if who != "":
    who = who.rstrip().lstrip()
else:
    print("Please choose someone to write for. Try again.")
    exit()

if site != "":
    
    site = site.rstrip().lstrip()

    if site == "chumba":
        site = "Chumba"
    elif site == "gp":
        site = "GP"
    elif site == "puls" or site == "pulsz":
        site = "Pulsz"
    elif site =="ll":
        site = "LL"
    elif site =="stake" or site == "Stake":
        site = "Stake"
else:
    site = "Pulsz"

if write_on != "":
    write_on = write_on.rstrip().lstrip()
    if write_on == "E" or write_on == "e" or write_on == "envelope" or write_on == "Envelope" or write_on == "Envelopes" or write_on == "envelopes":
        write_on = "Envelopes"
    else:
        write_on = "Inserts"
else:
    write_on = "Envelopes"

plot_path += who+"/"+site+"/"+write_on+"/"
print("plot path: "+plot_path)

# Setting up AxiDraw System
_, _, files = next(os.walk(plot_path))
template_count = len(files)

print("Writing for: "+plot_path + "; for site: "+site+"; on "+write_on+"; with file count "+str(template_count))

plotter_ports = []

if __name__ == "__main__":
    for port in list_ports.comports():
        if "USB" in port.hwid:
            print(f"Name: {port.name}")
            print(f"Description: {port.description}")
            #print(f"Location: {port.location}")
            #print(f"Product: {port.product}")
            print(f"Manufacturer: {port.manufacturer}")
            print(f"ID: {port.pid}")

            if(port.manufacturer == "Microsoft"):
                plotter_ports.append(port.name)

#get this weeks count for the template
row_to_pull = who + ' - ' + site+ ' - '+ write_on

#print("row to pull: " + row_to_pull)
#print(p_df)
pulled_row = p_df.query("Type == @row_to_pull")
current_week_value = pulled_row['Current Week'].iloc[0]
print("Current Weeks Value: "+str(current_week_value))

column = tracker_worksheet.find("Current Week")
row = tracker_worksheet.find(row_to_pull)
column = string.ascii_uppercase[column.col - 1] 
row = row.row
cell = column+str(row)

#print("cell is : "+ cell)

today = date.today()

# get todays date and corresponding cell value
year = today.strftime("%Y")
abv_year = year[-2:]
today_date = today.strftime("%m/%d/")+abv_year
#print("Todays date: ", today_date)

todays_value = pulled_row[str(today_date)].iloc[0]
today_col = tracker_worksheet.find(str(today_date))
today_col = string.ascii_uppercase[today_col.col - 1]
today_cell = today_col+str(row)

if todays_value == "":
    todays_value = 0

print("Todays cell value: "+str(todays_value))

start_at = 1

if current_week_value == "":
    current_week_value = 0

#determine which templates to plot
if current_week_value > template_count:
    start_at = current_week_value % template_count
else:
    start_at = current_week_value

# Functions

# Return to home position and clear thread information
def home():
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.flush_cmd()
    swift.set_wrist(90)

# Pickup an envelope from stack 1
def pickup1(z_change):
    swift.set_position(x=6, y=-215, z=150, speed=100, wait=True)
    swift.set_position(x=6,y=-230, speed=10000)
    swift.set_wrist(90)
    swift.set_pump(True)
    time.sleep(0.5)
    
    y_position = -245
    z_position = 60 - z_change

    swift.set_position(x=6, y=y_position, z=z_position, speed=100, wait=True) # find the incremental z rate change
    time.sleep(0.75)
    swift.set_position(x=6, y=y_position, z=180, speed=100, wait=True)
    time.sleep(1)
    home()

# Place an envelope that has already been picked up using pickup1 at plotter 1
def place1():
    swift.set_wrist(54)
    swift.set_position(x=229, y=-30, z=-35, speed=100, wait=True)
    time.sleep(0.5)
    swift.set_position(x=315, y=-140, z=-45, speed=100, wait=True)
    time.sleep(0.5)
    swift.set_pump(False)

# Pickup an envelope from stack 2
def pickup2(z_change):
    swift.set_position(x=6, y=215, z=150, speed=100, wait=True)
    swift.set_wrist(90)
    swift.set_pump(True)
    time.sleep(0.5)

    y_position = 265
    z_position = 60 - z_change
    
    swift.set_position(x=6, y=y_position, z=z_position, speed=100, wait=True) # find the incremental z rate change
    time.sleep(0.75)
    swift.set_position(x=6, y=y_position, z=180, speed=100, wait=True)
    time.sleep(1)
    home()

# Place an envelope that has already been picked up using pickup2 at plotter 2
def place2():
    swift.set_wrist(85)
    # swift.set_position(x=187, y=97, z=-35, speed=100, wait=True)
    swift.set_position(x=250, y=0, z=-35, speed=100, wait=True)
    time.sleep(0.5)
    swift.set_position(x=270, y=170, z=-35, speed=100, wait=True)
    time.sleep(0.5)
    swift.set_pump(False)
    home()

# Remove envelope from plot 1 and place into box 1
def removedrop1():
    home()
    swift.set_wrist(54)
    swift.set_position(x=293, y=-110, z=-65, speed=100, wait=True) # change the z value only
    time.sleep(0.5)
    swift.set_pump(True)
    time.sleep(0.5)
    swift.set_position(z=-45)
    time.sleep(0.5)
    swift.set_position(x=300, y=-105, z=-45, speed=100, wait=True)
    swift.set_position(x=229, y=-30, z=-35, speed=100, wait=True)
    home()
    swift.set_position(x=50, y=-180, z=180, speed=1000, wait=True)

    swift.set_wrist(90)

    #drop off position
    swift.set_position(x=50, y=-315, z=140, speed=1000, wait=True)
    time.sleep(0.5)
    swift.set_pump(False)
    time.sleep(0.5)
    swift.set_position(x=50, y=-180, z=180, speed=1000, wait=True)
    swift.set_position(x=250, y=0, z=180, speed=1000, wait=True)
    swift.set_wrist(90)


# Remove envelope from plot 2 and place into box 2
def removedrop2():
    home()
    swift.set_wrist(94)
    swift.set_position(x=240, y=170, z=-65, speed=1000, wait=True) # change the z value only
    time.sleep(0.5)
    swift.set_pump(True)
    time.sleep(0.5)
    swift.set_position(z=-45)
    time.sleep(0.5)
    swift.set_position(x=260, y=140, z=-45, speed=1000, wait=True)
    # swift.set_position(x=260, y=140, z=-45, speed=1000, wait=True)
    # swift.set_position(x=187, y=97, z=-35, speed=100, wait=True)
    # swift.set_position(x=229, y=-30, z=-35, speed=100, wait=True)
    home()
    swift.set_position(x=50, y=180, z=180, speed=1000, wait=True)

    swift.set_wrist(110)

    #drop off position
    swift.set_position(x=50, y=315, z=140, speed=1000, wait=True)
    time.sleep(0.5)
    swift.set_pump(False)
    time.sleep(0.5)
    swift.set_position(x=50, y=180, z=180, speed=1000, wait=True)
    swift.set_position(x=250, y=0, z=180, speed=1000, wait=True)
    swift.set_wrist(90)

# AxiDraw Stuff (Still have to update comments on axidraw code)
def GoPlot(port, run):
    #print("Printing to port "+ port)

    last_letter = write_on[-1]
    singular_type = write_on

    #run = run -1 


    if last_letter == "s":
       singular_type = write_on.rstrip(write_on[-1])

    who2 = who.replace(" ", "!")

    file_name = who + "-" + site + "-" + singular_type +""+ str(run) +".svg"
    #print("The chosen file is "+ file_name)

    path2 = plot_path.replace(" ", "!")

    file_to_print = path2 + file_name
    print("Print to port: " +port + " with template named: " +file_to_print)

    os.system(f"python axi_automation_test2.py {port} {file_to_print}")


def continuous():
    global count
    global current_count
    for x in range(235): # change the number in the range function to change how many times it runs through
        count += 1
        # Setting up plot 1
        # home position and clear
        home()
        # # Pick up from stack 1
        pickup1(0.42*count)
        # # Place at plotter 1
        place1()
        
        # Setting up plot 2 
        # home position and clear
        home()
        # Pick up from stack 1
        pickup2(0.42*count)
        # Place at plotter 1
        place2()

        # Begin AxiDraw plotting
        current_run = start_at
        for plotter in plotter_ports:
            #print("Current Run: "+str(current_run)+ "; and template count is: "+ str(template_count))
            if current_run > template_count:
                current_run = 1
                current_count -= 1
            elif current_run != template_count:
                current_run += 1
                current_count -= 1

            threading.Thread(target = GoPlot, args= (plotter,current_run,)).start()
            if current_run == template_count:
                current_run += 1
                current_count -= 1

        # # Return to home and wait for AxiDraw to finish plotting
        home()
        time.sleep(90)

        # # Remove envelopes and place into boxes
        removedrop1()
        removedrop2()

        
        time.sleep(0.5)
        swift.set_pump(False)
        home()

# Run full setup
continuous()
# swift.set_position(z=-65)
exit()