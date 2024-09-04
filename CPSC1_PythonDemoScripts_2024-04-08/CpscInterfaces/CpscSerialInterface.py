###############################################################################
# File name:      CpscSerialInterface.py
# Creation:       2021-12-10 16:53:09
# Updated:        2022-11-01
# Author:         JPE, Robin Drossaert
# Python version: 3.9
#
# CpscSerialInterface handles serial communication with a CPSC.
#
# The port name and baudrate are variable.
# Other communication settings are fixed as follows:
# - Data bits = 8 bits
# - Parity = none
# - Stop bits = 1
# - Timeout = 10 seconds
# - Flow control = none
#
# Since the CPSC always send a response following a command, only the
# WriteRead() function should be used for normal communication.
#
# Added functionality: option to automatically add \r\n to a txMessage and
# to remove \r\n from an rxMessage (txTermination argument)
#
# Use the CpscSerialInterface in a 'with' 'as' construction to ensure
# the port is always closed after running the program.
###############################################################################

# 3rd party imports
import serial

class CpscSerialInterface:

    def __init__(self,comPort,baudrate):
        self.com = serial.Serial(port = comPort,
                                 baudrate = baudrate,
                                 bytesize = 8,
                                 parity = serial.PARITY_NONE,
                                 stopbits = serial.STOPBITS_ONE,
                                 timeout = 10,
                                 xonxoff = False)
        self.com.flushInput()

    def __enter__(self):
        return self

    def __exit__(self,exType,exValue,trcbck):
        self.Close()

    def Close(self):
        self.com.close()

    def Write(self,txMessage):
        self.com.write(txMessage.encode('ascii')) # Sent message as ASCII string

    def Read(self):
        rxMessage =  self.com.read_until(b'\r\n') # Read until termination characters
        return rxMessage.decode() # Convert message to Python3 string

    def WriteRead(self, txMessage, txTermination):
        if txTermination == 0:
            self.Write(txMessage)
            rxMessage = self.Read()
            return rxMessage
        else:
            self.Write(txMessage + '\r\n')
            rxMessage = self.Read()
            rxMessageClean = rxMessage.replace('\r\n', '')
            return rxMessageClean
