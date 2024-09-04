###############################################################################
# File name:      CpscEthernetInterface.py
# Creation:       2021-12-13 12:47:32
# Author:         JPE, Robin Drossaert
# Python version: 3.9.7
#
# CpscEthernetInterface handles Ethernet (TCP) communication with a CPSC.
#
# The IP address is variable.
# Other communication settings are fixed as follows:
# - Port = 2000
# - Timeout = 10 seconds
#
# Since the CPSC always send a response following a command, only the
# WriteRead() function should be used for normal communication.
#
# Use the CpscEthernetInterface in a 'with' 'as' construction to ensure
# the socket is always closed after running the program.
###############################################################################

# 3rd party imports
import socket
import time

class CpscEthernetInterface:

    def __init__(self,ipAddress,tcpPort):
        self.port = tcpPort
        self.timeout = 10
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((ipAddress,self.port))
        self.sock.setblocking(False)

    def __enter__(self):
        return self

    def __exit__(self,exType,exValue,trcbck):
        self.Close()

    def Close(self):
        self.sock.close()

    def Write(self,txMessage):
        self.sock.send(txMessage.encode('ascii')) # Sent message as ASCII string

    def Read(self):
        rxMessage = ''.encode()
        startTime = time.time()
        passedTime = 0
        # Repeat read until termination characters or nothing has been received for 10 seconds
        while rxMessage.find('\r\n'.encode()) == -1 and passedTime < self.timeout:
            try:
                rxMessage += self.sock.recv(100)
                startTime = time.time() # Byte received so reset timer
            except:
                passedTime = time.time() - startTime
        return rxMessage.decode() # Convert message to Python3 string

    def WriteRead(self, txMessage):
        self.Write(txMessage)
        rxMessage = self.Read()
        return rxMessage
