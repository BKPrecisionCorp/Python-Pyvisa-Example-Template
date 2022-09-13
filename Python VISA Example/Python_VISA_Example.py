# Python for Test and Measurement

# Requires VISA installed on controlling PC
# 'https://www.ni.com/en-us/support/downloads/drivers/download.ni-visa.html#409839'
#
# Requires PyVISA to use VISA in Python
# 'https://pypi.org/project/PyVISA/'

# Python 3.9.7
# pyvisa 1.11.x
# XlsxWriter 3.0.1
#datetime
# pyusb 1.1.1
##"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
## Copyright © 2021 B&K Precision Corporation. All rights reserved.
##
## You have a royalty-free right to use, modify, reproduce and distribute this
## example files (and/or any modified version) in any way you find useful, provided
## that you agree that B&K Precision has no warranty, obligations or liability for any
## Sample Application Files.
##
##"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

# Example Description:  
# A python template utilizing pyvisa in order to connect and control the Bk precision instruments  
# SCPI commands must be edited based on the instrument to be connected
# The application performs the following:
#   Import the pyvisa libraries;
#   Establish a visa resource manager;
#   Open a connection to the specified instrument's VISA address as acquired 
#   Via a loop the following: 
#       Acquire data based on the query entered.
#       Logs collected data to an excel file



import pyvisa
import csv
import time
from datetime import datetime
import xlsxwriter
from time import sleep
from pyvisa import ResourceManager, constants

####____INITIALIZE INSTRUMENT____####
def init():

    global IDN
    global IDN_list
    global src

    rm = pyvisa.ResourceManager()
    li = rm.list_resources()
    choice = ""

    while (choice == ""):
        for index in range (len(li)):
            print(str(index) + "-" + li[index])
        choice = input ("Select DUT:")  
        try:
            if(int(choice) > len(li) -1 or int (choice) < 0):
                choice = ""
                print("Invalid Input\n")
        except:
            print("Invalid Input\n")
            choice = ""

    ####____OPEN SESSION____####
    src = rm.open_resource(li[int(choice)])
    src.timeout = 10000 
    
    ####____Query ID____####
    IDN = src.query("*IDN?\n")
    print(IDN)
    IDN_list = IDN.split(",")               #index = [manufacture, model, SN, FW]


    return 0


####____Handler Function____####
def handle_event(resource, event, user_handle):
    resource.called = True
    print(f"Handled event {event.event_type} on {resource}")

####____CREATE EXCEL FILE____####
def createworkbook():
    ####____VARIABLES____####
    Date_Time = (datetime.now())
    Time = str(Date_Time.time())
    Time1 = Time.replace(":","_")
    Date = str(Date_Time.date())



    ####____CREATE WORKBOOK____####
    outWorkbook  = xlsxwriter.Workbook(IDN_list[1] + "-" + Date + "_" + Time1 + ".xlsx") 
    Sheet1 = outWorkbook.add_worksheet("Sheet1")  
    cell_format = outWorkbook.add_format() 
    cell_format.set_num_format("0.000")


    ####____WORKBOOK FORMATS____####

    num_format = outWorkbook.add_format({  
        'bold': 0,  
        'border': 0,  
        'align': 'center',  
        'valign': 'vcenter',
        'num_format': '0.000'}) 

    merge_format = outWorkbook.add_format({  
        'bold': 1,  
        'border': 1,  
        'align': 'center',  
        'valign': 'vcenter'}) 

    centerbold_format = outWorkbook.add_format({  
        'bold': 1,  
        'border': 0,  
        'align': 'center',  
        'valign': 'vcenter'}) 

    #increase column A nd B width
    Sheet1.set_column(0, 1, 16) 

    # Merge 6 columns for ID  
    Sheet1.merge_range('A1:F1', IDN, merge_format)  
  
    #Basic information of acquired data  
    Sheet1.write('A3', 'Date and Time', centerbold_format)  
    Sheet1.merge_range('B3:D3', Date + "_" + Time,centerbold_format)  
    Sheet1.write('A5', "Measurement", centerbold_format) 
    Sheet1.merge_range('B5:D5', "Time", centerbold_format)  

    #variables for loop 
    i = 7 

    try:
        while True:
            dt_obj2 = datetime.now()
            getimestamp = dt_obj2.strftime("%H:%M:%S")
        
            # Edit query for the data you want to collect (currently querying data of power supply with multiple channels like the 9140)
            meas = "%s"%(src.query("MEAS:ALL?\n"))
            meas1 = meas.replace("\n","")
            meas_li = (meas1.split(","))
            print(float(meas_li[0]))

            Meas = "A%d"%i
            timestamp = "B%d:D%d"%(i,i)
       

            Sheet1.write(Meas, float(meas_li[0]), num_format)
            Sheet1.merge_range(timestamp, getimestamp, centerbold_format)
        
            i += 1

            time.sleep(.050)


    except KeyboardInterrupt:                               # CTRL + C interrupt
        outWorkbook.close()
        return print("Data Acquired")


####____Set Parameters____####
def parameters():
    
    #Edit for the type of instrument you are using currently set for power supplies
    MIN_voltage = src.query("VOLT:MIN")
    MAX_voltage = src.query("VOLT:MAX?")

    print("Please enter a voltage between " + MIN_voltage + " and " + MAX_voltage)
    voltage = input()
    src.write("VOLT " +voltage)
    time.sleep(.025)

    MIN_current = src.query("CURR:MIN?")
    MAX_current = src.query("CURR:MAX?")

    print("Please enter a current between " + MIN_current  + " and " + MAX_current )
    currrent = input()
    src.write("CURR " + currrent)
    time.sleep(.025)

    return 0



def main():

    init()

    #Event Handler
    src.called = False

    # Type of event we want to be notified about
    event_type = constants.EventType.service_request
    # Mechanism by which we want to be notified
    event_mech = constants.EventMechanism.queue

    wrapped = src.wrap_handler(handle_event)

    user_handle = src.install_handler(event_type, wrapped, 42)
    src.enable_event(event_type, event_mech, None)

    # Instrument specific code to enable service request
    # (for example on operation complete OPC)
    src.write("*SRE 1")
    src.write("INIT")

    while not src.called:
        sleep(10)

    parameters()

    createworkbook()

    src.close()
    print("Application Complete")

    src.disable_event(event_type, event_mech)
    src.uninstall_handler(event_type, wrapped, user_handle)


if __name__ == '__main__': 
 proc = main()




