#Nathen Feldgus

#Displays output of EMMS meter


import serial           # For overall communication between  the computer and devices

def main():


    ekm = serial.Serial(
        port='COM4',
        baudrate=9600,
        parity=serial.PARITY_EVEN,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.SEVENBITS,
        xonxoff=0,
        timeout=4       # time before reading
    )

    emms = serial.Serial(
        port='COM5',
        baudrate=9600,
        parity=serial.PARITY_EVEN,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.SEVENBITS,
        xonxoff=0,
        timeout=4          # constant so the emms is reading within a short time of the EKM
    )
    emmsResponse = emms.read(10000)
    print(emmsResponse)

    meterNum = "000010002935"
    
    ekm.write(str.encode("\x2F\x3F"+meterNum+"\x30\x30\x21\x0D\x0A"))       # Send Request A to v4 Meter
    ekmResponse = ekm.read(10000).decode("utf-8")

    print(ekmResponse)

main()
