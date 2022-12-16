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

from pickle import FALSE
import time
from uarm import swift
from uarm.wrapper.swift_api import SwiftAPI

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

# Initializing dict variables for uArm robot
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

# Get and Print uArm information
if sys.argv[2] is not None:
    robo_port = pod_dict[sys.argv[2]]["robot"]

accessed = False
while not accessed:
    try:
        #swift = SwiftAPI(filters={'hwid': 'USB VID:PID=2341:0042'})
        swift = SwiftAPI(port=robo_port)
        swift.waiting_ready(timeout=3)
        accessed = True
    except:
        print(str(e))
        time.sleep(0.2)

print('device info: ')
print(swift.get_device_info())
print(swift.port)

# Print Envelope Information
current_count = int(sys.argv[3])

stack = "one"

if current_count > 300:
    stack = "two"

# Initialize variables for AxiDraw System and print path directory for svg plot files
if sys.argv[4] is not None:
    who = sys.argv[4]
    who = who.replace("-", " ")
    write_on = sys.argv[5]
    site = sys.argv[6]
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

# Define methods for uArm robot
def set_new_insert():
    print("set new insert")

def remove_insert():
    print("remove insert")

def remove_envelope():
    print("remove envelope")

def set_new_envelope():
    print("set new envelope")

def remove_and_set_new_envelope():
    swift.set_position(x=250, y=0, z=50, speed=1000, wait=True)

def pickup_new_envelope_stack_one(next_step):
    global swift
    global current_count

    print(current_count)
    y_rate_change = 0.1
    z_rate_change = 0.57
    z_start = 108
    y_start = -215

    #Tested with a mix of envelope types
    if current_count > 25 and current_count < 100:
        z_start = 103
        y_start = -215
    elif current_count > 99 and current_count < 150:
        y_rate_change = 0.15
        z_rate_change = 0.54
    elif current_count > 149 and current_count < 175:
        y_rate_change = 0.1
        z_rate_change = 0.54
    elif current_count > 174 and current_count < 200:
        y_rate_change = 0.15
        z_rate_change = 0.52
    elif current_count > 199 and current_count < 290:
        y_rate_change = 0.1
        z_rate_change = 0.54
    elif current_count > 289:
        y_rate_change = 0.1
        z_rate_change = 0.58

    #Determine the variable envelope pickup position
    new_envelope_position = z_start - int(float(current_count) * z_rate_change)
    print("new_envelope_position: " + str(new_envelope_position))
    # round to two decimals
    new_envelope_position = round(new_envelope_position, 2)

    #At z of 20, y needs to be 180
    y_pos = y_start + int(float(current_count)*y_rate_change)
    print("y_pos: " + str(y_pos))
    y_pos = round(y_pos, 2)


    #home position and clear
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.flush_cmd()

    swift.set_position(x=6, y=-215, z=150, speed=100, wait=True)

    swift.set_wrist(90)

    #adjusting y position for decreasing envelope stack
    swift.set_pump(True)
    time.sleep(0.5)
    swift.set_position(x=6, y=y_pos, z=new_envelope_position, speed=100, wait=True)
    swift.set_position(x=6, y=y_pos, z=(new_envelope_position -2), speed=100, wait=True)
    time.sleep(0.75)
    swift.set_position(x=6, y=y_pos, z=150, speed=100, wait=True)
    time.sleep(1)
    swift.set_position(x=3.5, y=-230, z=150, speed=1000, wait=True)
    
    if next_step == 'hold':
        time.sleep(5)
        swift.set_pump(False)

def pickup_new_envelope_stack_two(next_step):
    global swift
    global current_count

    #s_count = current_count - 300
    s_count = current_count

    y_rate_change = 0.1
    z_rate_change = 0.57
    z_start = 90
    y_start = 215

    #Tested with a mix of envelope types
    if s_count > 25 and s_count < 100:
        z_start = 103
        y_start = 215
    elif s_count > 99 and s_count < 150:
        y_rate_change = 0.15
        z_rate_change = 0.54
    elif s_count > 149 and s_count < 175:
        y_rate_change = 0.1
        z_rate_change = 0.54
    elif s_count > 174 and s_count < 200:
        y_rate_change = 0.15
        z_rate_change = 0.52
    elif s_count > 199 and s_count < 290:
        y_rate_change = 0.1
        z_rate_change = 0.54
    elif s_count > 289:
        y_rate_change = 0.1
        z_rate_change = 0.58

    #Determine the variable envelope pickup position
    new_envelope_position = z_start - int(float(s_count) * z_rate_change)
    print("new_envelope_position: " + str(new_envelope_position))
    # round to two decimals
    new_envelope_position = round(new_envelope_position, 2)

    #At z of 20, y needs to be 180
    y_pos = y_start - int(float(s_count)*y_rate_change)
    print("y_pos: " + str(y_pos))
    y_pos = round(y_pos, 2)


    #home position and clear
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.flush_cmd()

    swift.set_position(x=6, y=215, z=150, speed=1000, wait=True)

    swift.set_wrist(90)

    #adjusting y position for decreasing envelope stack
    swift.set_pump(True)
    time.sleep(0.5)
    swift.set_position(x=6, y=y_pos, z=new_envelope_position, speed=100, wait=True)
    swift.set_position(x=6, y=y_pos, z=(new_envelope_position -2), speed=100, wait=True)
    time.sleep(0.75)
    swift.set_position(x=6, y=y_pos, z=150, speed=100, wait=True)
    time.sleep(1)
    swift.set_position(x=3.5, y=230, z=150, speed=1000, wait=True)

    if next_step == 'hold':
        time.sleep(5)
        swift.set_pump(False)

#This is going to be position 2
def place_position_one_envelope(next_step):
    global swift
    global current_count
    global stack

    #home position
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.set_wrist(54)
    swift.set_position(x=229, y=-30, z=-35, speed=100, wait=True)
    time.sleep(0.5)

    #Because we grab at different heights on envelope, we need to adjust the y position based on count
    #Higher the count, the less we need to move the y position
    temp_count = current_count

    if stack == "one":

        temp_count = current_count - 300
    
        x_pos = 295 - float(temp_count)*0.05
    
        swift.set_position(x=x_pos, y=-100, z=-45, speed=100, wait=True)
    else:

        x_pos = 287 - float(temp_count)*0.05
        swift.set_wrist(50)
    
        swift.set_position(x=x_pos, y=-119, z=-45, speed=100, wait=True)
    
    time.sleep(0.5)
    swift.set_pump(False)
    time.sleep(0.5)
    if next_step != "hold":
        swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)


#This is going to be position 2
def pickup_position_one_envelope():
    global swift
    global current_count

    #home position
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.set_wrist(50)
    swift.set_position(x=293, y=-110, z=-65, speed=100, wait=True)
    time.sleep(0.5)
    swift.set_pump(True)
    time.sleep(0.5)
    swift.set_position(x=300, y=-105, z=-45, speed=100, wait=True)
    swift.set_position(x=229, y=-30, z=-35, speed=100, wait=True)
    swift.set_position(x=250, y=0, z=150, speed=100, wait=True)
    swift.set_wrist(90)

#Right side position
def place_position_two_envelope(next_step):
    global swift
    global current_count

    #home position
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.set_wrist(94)
    swift.set_position(x=187, y=97, z=-35, speed=100, wait=True)
    time.sleep(0.5)

    #Because we grab at different heights on envelope, we need to adjust the y position based on count
    #Higher the count, the less we need to move the y position
    #y_pos = 163 - float(current_count)*0.08
    
    #swift.set_position(x=270, y=y_pos, z=-40, speed=100, wait=True)

    temp_count = current_count
    if stack == "one":
        temp_count = current_count - 300
    
        x_pos = 250 - float(temp_count)*0.05
    
        swift.set_position(x=x_pos, y=168, z=-45, speed=100, wait=True)
    else:
        x_pos = 275 - float(temp_count)*0.05
    
        swift.set_position(x=x_pos, y=137, z=-45, speed=100, wait=True)
    
    time.sleep(0.5)
    swift.set_pump(False)
    time.sleep(0.5)
    if next_step != "hold":
        swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
        


#This is going to be position 2
def pickup_position_two_envelope():
    global swift
    global current_count

    #home position
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.set_wrist(94)
    swift.set_position(x=260, y=150, z=-70, speed=1000, wait=True)
    time.sleep(0.5)
    swift.set_pump(True)
    time.sleep(0.5)
    swift.set_position(x=260, y=150, z=-45, speed=1000, wait=True)
    swift.set_position(x=187, y=97, z=-35, speed=100, wait=True)
    swift.set_position(x=250, y=0, z=150, speed=100, wait=True)
    swift.set_wrist(90)


def drop_complete_from_home_envelope_stack_one():
    global swift

    #home position and clear
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)

    swift.set_position(x=50, y=-180, z=150, speed=1000, wait=True)

    swift.set_wrist(65)

    #drop off position
    swift.set_position(x=50, y=-336, z=120, speed=1000, wait=True)
    time.sleep(0.5)
    swift.set_pump(False)
    time.sleep(0.5)
    swift.set_position(x=50, y=-180, z=150, speed=1000, wait=True)
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.set_wrist(90)

def drop_complete_from_home_envelope_stack_two():
    global swift

    #home position and clear
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)

    swift.set_position(x=50, y=180, z=150, speed=1000, wait=True)

    swift.set_wrist(115)

    #drop off position
    swift.set_position(x=50, y=336, z=120, speed=1000, wait=True)
    time.sleep(0.5)
    swift.set_pump(False)
    time.sleep(0.5)
    swift.set_position(x=50, y=180, z=150, speed=1000, wait=True)
    swift.set_position(x=250, y=0, z=150, speed=1000, wait=True)
    swift.set_wrist(90)

#print("In run robo")
#print("The place is: " + sys.argv[1])
#print("The type is: " + sys.argv[2])
print("The count is: " + str(sys.argv[4]))

def get_new_envelope():
    if stack == 'two':
        pickup_new_envelope_stack_two()
    else:
        pickup_new_envelope_stack_one()

def drop_complete_envelope():
    if stack == 'two':
        drop_complete_from_home_envelope_stack_two()
    else:
        drop_complete_from_home_envelope_stack_one()

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

if sys.argv[1] == "place":
    if sys.argv[2] == "Envelope":
        get_new_envelope()
        time.sleep(1)
        place_position_one_envelope()
        #get_new_envelope()
        time.sleep(1)
        #place_position_two_envelope()
        swift.flush_cmd()
        exit()
    else:
        set_new_insert()
elif sys.argv[1] == "remove":
    if sys.argv[2] == "Envelope":
        print("removing envelopes")
        pickup_position_one_envelope()
        time.sleep(1)
        drop_complete_envelope()
        time.sleep(1)
        pickup_position_two_envelope()
        time.sleep(1)
        drop_complete_envelope()
        time.sleep(1)
        get_new_envelope()
        time.sleep(1)
        place_position_one_envelope()
        get_new_envelope()
        time.sleep(1)
        place_position_two_envelope()
        exit()
    else:
        remove_insert()
elif sys.argv[1] == "pick1":
    stack = "one"
    pickup_new_envelope_stack_one("hold")
    exit()
elif sys.argv[1] == "pick2":
    stack = "two"
    pickup_new_envelope_stack_two("hold")
    exit()
elif sys.argv[1] == "place1":
    place_position_one_envelope()
    exit()
elif sys.argv[1] == "place1hold":
    place_position_one_envelope("hold")
    exit()
elif sys.argv[1] == "place2":
    place_position_two_envelope()
    exit()
elif sys.argv[1] == "place2hold":
    place_position_two_envelope("hold")
    exit()
elif sys.argv[1] == "pick1place1":
    stack = "one"
    pickup_new_envelope_stack_one("1")
    time.sleep(1)
    place_position_one_envelope("1")
    exit()
elif sys.argv[1] == "pick2place2":
    stack = "two"
    pickup_new_envelope_stack_two("1")
    time.sleep(1)
    place_position_two_envelope("1")
    exit()
elif sys.argv[1] == "remove1drop1":
    stack = "one"
    pickup_position_one_envelope()
    time.sleep(1)
    drop_complete_from_home_envelope_stack_one()
    exit()
elif sys.argv[1] == "remove2drop2":
    stack = "two"
    pickup_position_two_envelope()
    time.sleep(1)
    drop_complete_from_home_envelope_stack_two()
    exit()
elif sys.argv[1] == "continuous":
    stack = "one"
    pickup_new_envelope_stack_one("1")
    # time.sleep(1)
    # place_position_one_envelope("1")
    # stack = "two"
    # pickup_new_envelope_stack_two("1")
    # time.sleep(1)
    # place_position_two_envelope("1")
    # stack = "one"
    # pickup_position_one_envelope()
    # time.sleep(1)
    # drop_complete_from_home_envelope_stack_one()
    # stack = "two"
    # pickup_position_two_envelope()
    # time.sleep(1)
    # drop_complete_from_home_envelope_stack_two()




#Y = -355 to 355
#X = 50 to 310
#Z = -100 to 150mm