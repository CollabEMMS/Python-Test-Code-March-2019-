#Nathen Feldgus
#Ben Weaver

#Logs the following data:

#DATE TIME,
#(last logged) Power (EKM), Energy (EKM), Age (EKM), 
#(last logged) Power (EMMS), Energy (EMMS), Age (EMMS)
#Log interval (seconds)
#Age of data
#+anything else helpful
#Sput out allotment



import serial           # For overall communication between  the computer and devices
import datetime         # For getting the date of the test
import time             # For getting the system timestamps
import pyexcel          # For creating the .csv file
import os               # For file directory information



# Gets initial data for values that we want a change (time, energies)

def firstContact(meterNum):
    emmsEnergyInitial = int(0)
    ekmEnergyInitial = int(0)    #Set defaults here in case something funky happens


    emmsResponse = emms.read(10000)                                         # Read emms
    packets = []                                                            # 
    packets += emmsResponse.decode("utf-8").split('%')                      # Dump the responses into an array
    if (packets == [""]):                                                   # If the packets are empty, the EMMS is not connected
        print("\nERROR!  No EMMS response!")
        print("\n\nEMMS Initial Energy: 0 \n")                
        #testInfo = testInfo + "\nNo EMMS response at startup"
    else:
        for i in range(1, len(packets)):

            emmsRead = packets[i].split(' ')                                    # Split the individual entries' type and values
            
            if (emmsRead[1] == 'Power'):                                        # Look for a packet that is labeled "Power"
                emmsEnergyInitial = int(emmsRead[2].split(':')[1])              # Parse out just the energy value (and print)
                print("EMMS Initial Energy: " + str(emmsEnergyInitial) + "\n")                
                break

    ekm.write(str.encode("\x2F\x3F"+meterNum+"\x30\x30\x21\x0D\x0A"))       # Send Request A to v4 Meter
    ekmResponse = ekm.read(10000).decode("utf-8")                           # Get response and decode
    
    if (ekmResponse[20:24] == ''):                                          # Parse for energy value
        print("ERROR!  No EKM response!")
        #testInfo = testInfo + "\nNo EKM response at startup" - I want a comment that the EKM meter isn't connected but this is throwing up an error when I try it
    else:
        ekmEnergyInitial =  int(ekmResponse[20:24])                             
        
    print("\n\nEKM Initial Energy: " + str(ekmEnergyInitial))               # Print (for user feedback)

            
    return (ekmEnergyInitial, emmsEnergyInitial)         # Return the values for use in loop function            

    # Send close to meter
    ekm.write(str.encode("\x01\x42\x30\x03\x75"))
    emms.write(str.encode("\x01\x42\x30\x03\x75"))

    # Close connection to serial port
    ekm.close
    emms.close





# Communicates with the meter based on the user-defined time between readings.
# Within the loop is also the excel file creation code.  On retrival, values are
# dumped into the excel file in a new line, and the loop repeats


def readMetersLoop(meterNum, osStartTime, ekmEnergyInitial, emmsEnergyInitial):

    while (True):

        # READ EMMS METER

        emmsResponse = emms.read(10000)                                 #
        packets = []                                                    #
        packets += emmsResponse.decode("utf-8").split('%')              # Record input into an array
        emmsPower = 0                                                   # Default values for emmsPower and emmsEnergyChange to prevent error if disconnected
        emmsEnergyChange = 0
        for i in range(1, len(packets)):

            emmsRead = packets[i].split(' ')
            if (len(emmsRead) > 1):
            
                if (emmsRead[1] == 'Power'):                                # Find a string that represents power
                    emmsPower = emmsRead[2].split(':')[0]
                    if (emmsPower != ''):
                        emmsPower = int(emmsPower)                          # Record Power
                    print("EMMS Power: " + str(emmsPower))
                    
                    if (len(emmsRead[2].split(':')) > 1):
                        emmsEnergyChange = emmsRead[2].split(':')[1]
                        if (emmsEnergyChange != ''):
                            emmsEnergyChange = int(emmsEnergyChange) - emmsEnergyInitial  # Record change in energy
                        print("EMMS Energy Change: " + str(emmsEnergyChange))                
                        break

        # READ EKM METER
        
        ekm.write(str.encode("\x2F\x3F"+meterNum+"\x30\x30\x21\x0D\x0A"))       # Send Request A to v4 Meter
        ekmResponse = ekm.read(10000).decode("utf-8")                           # Get response and decode

        if (ekmResponse[126:130] == ''):                                        #Get the Power value
            ekmPower = 0
        else:
            ekmPower = int(ekmResponse[126:130])                      
        print("\nEKM Power:  " + str(ekmPower))                         

        if (ekmResponse[20:24] == ''):                                          # Get the Energy value and find the change in energy
            ekmEnergyChange = int(-ekmEnergyInitial)
        else:
            ekmEnergyChange = int(ekmResponse[20:24]) - ekmEnergyInitial    
        print("EKM Energy Change: " + str(ekmEnergyChange))             


        #LOG DATA
        
        excelSheet(name, osStartTime, ekmPower, emmsPower, ekmEnergyChange, emmsEnergyChange)


# Makes an entry in the excel sheet based on the input arguments

def excelSheet(name, osStartTime, ekmPower, emmsPower, ekmEnergyChange, emmsEnergyChange):


    # Find elapsed time from the computer's perspective
    osElapsedTime = round(time.time() - osStartTime, 1)
    #Convert osElapsedTime to hours
    hours = int(osElapsedTime//3600)
    minutes = int(((osElapsedTime)-(hours*3600))//60)
    seconds = int((osElapsedTime - (3600*hours) - (60*minutes)))
    timeData = ""
    if (hours != 0):
        timeData += str(hours) + "h, "
    if (minutes != 0):
        timeData += str(minutes) + "m, "
    timeData += str(seconds) + "s"
    
    # Adds a row with values
    data.append([timeData,ekmPower,emmsPower,ekmEnergyChange,emmsEnergyChange])

    # Saves the file
    save(data,name)




def save(data, name):
    try:
        pyexcel.save_as(array=data,
                        dest_file_name=name+".csv",
                        dest_delimiter=',')
    except:
        print("COULD NOT SAVE DATA, PLEASE CLOSE THE FILE!")
    else:
        print("Data saved \n\n")  # Feedback print statement

#check if file name already exists, if it does, rename the file
def fileCheck(name, filePath):
    if(os.path.isfile(filePath)):
        print("The specified file already exists")
        while(os.path.isfile(filePath)):
            name += " - Copy"
            filePath = os.path.abspath(name+".csv") 
        print("Renamed file to " + name)
    else:
        print("Creating new file...")
    return name, filePath

def main():
    #SETUP
    global date
    date = str(datetime.datetime.today()).split()[0]    # Gets the date in a certain format

    print("Default values - On Messiah ThinkCentre, EKM is on COM4 and EMMS on COM5")
    print("Default delay - 60 seconds (1 minute)")
    print("Default filename - today's date, " + date)

    
    ekmPort = input("EKM meter on port COM")
    #Integer cannot be '', so set to int after making sure it isn't blank
    if (ekmPort == ''):
        ekmPort = 4
        print("4")
    else:
        ekmPort = int(ekmPort)
    emmsPort = input("EMMS meter on port COM")
    if (emmsPort == ''):
        emmsPort = 5
        print("5")
    else:
        emmsPort = int(emmsPort)
        
    global delay

    #delay-seconds between reads
    delay = input("Time between readings (in seconds): ")
    if (delay == ''):
        delay = 60  #For testing purposes, the default delay will be 1 second
        print("60") #Please change this after this code is finished
    else:
        delay = int(delay)

    # Excel Sheet Setup
    global name
    global testInfo
    global filePath

    print("Default filename = today's date, " + date)
    name = str(input("Output Filename: "))
    if (name == ""):
        name = date
    testInfo = str(input("Testing Information: "))      # Allows user to put comments like testing parameters
    if (testInfo == ''):
        testInfo = "User had no comments"
    filePath = os.path.abspath(name+".csv")             # os.path.abspath finds the filepath for the new file
    name,filePath = fileCheck(name,filePath)

    global ekm
    global emms
    
    # Open connection to serial port
    # Creates serial objects for each of the meters
    ekm = serial.Serial(
        port='COM'+str(ekmPort),
        baudrate=9600,
        parity=serial.PARITY_EVEN,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.SEVENBITS,
        xonxoff=0,
        timeout=1       # constant so the emms is reading within a short time of the EMMS
    )
    emms = serial.Serial(
        port='COM'+str(emmsPort),
        baudrate=9600,
        parity=serial.PARITY_EVEN,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.SEVENBITS,
        xonxoff=0,
        timeout= delay          
    )

    # EKM Meter Number
    meterNum = "000010002935"


    # First Contact (gets the original energy info)
    ekmEnergyInitial, emmsEnergyInitial = firstContact(meterNum)
    osStartTime = time.time()

    #parse timestamp
    formattedStartTime = str(time.localtime(osStartTime)).split(", ")
    formattedStartTime = formattedStartTime[3].split("=")[1]+":"+formattedStartTime[4].split("=")[1]+":"+formattedStartTime[5].split("=")[1]



    # Get meter responses and log them in the spreadsheet
    
    global data     # Data is the spreadsheet's variable. The following info is for the table heading
    data = [["File Name",name],["Date",date],["Time",formattedStartTime],["Test Info",testInfo],["Directory",filePath],[],["System TIMESTAMP", "EKMPOWER", "EMMS POWER", "EKM Energy Used", "EMMS Energy Used"]]

    readMetersLoop(meterNum, osStartTime, ekmEnergyInitial, emmsEnergyInitial)
            
            

    # Send close to meter
    ekm.write(str.encode("\x01\x42\x30\x03\x75"))
    emms.write(str.encode("\x01\x42\x30\x03\x75"))

    # Close connection to serial port
    ekm.close
    emms.close

    print(filePath) # So user can find the file even if they just run the code in the command prompt

   
main()



# TODO

# Open Ended readings
# 5 minutes of EMMS meter
# Check end of EMMS string
# Non-blocking
