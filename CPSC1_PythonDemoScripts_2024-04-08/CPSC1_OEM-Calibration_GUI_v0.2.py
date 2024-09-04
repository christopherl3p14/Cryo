###############################################################################
# File name:      CPSC1_OEM-Calibration_GUI_vX.y.py
# Creation:       2022-12-07
# Last Update:    2023-01-31
# Author:         JPE 
# Python version: 3.9, requires pyserial and tkinter
# Description:    MVP GUI calibrating an OEM module for use with connected
#                 Cryo Optical Encoders (-COE)
# Disclaimer:     This program is provided 'As Is' without any express or 
#                 implied warranty of any kind.
###############################################################################

verNumber = 'v0.2'

# 3rd party imports
import time
import tkinter as tk
import subprocess as sp
import csv
from matplotlib import pyplot as plt
import threading as thrd

# JPE imports
from CpscInterfaces import CpscSerialInterface as CpscSerial

# Create GUI window
window = tk.Tk()

# Setup icon global window parameters
p1 = tk.PhotoImage(file = 'jpe.png')
window.iconphoto(False, p1)
window.title(f'CPSC1 OEM Calibration ({verNumber})') 
window.option_add( "*font", "Helvetica 11" )

# Setup frames within window
frm1 = tk.Frame(width=500, padx=10, pady=10)
frm2 = tk.Frame(width=500, padx=10, pady=10)
frm3 = tk.Frame(width=500, height=350, padx=10, pady=10) 
frm4 = tk.Frame(width=500, padx=10, pady=10) 
frm5 = tk.Frame(width=500, padx=10, pady=10) 
frm1.pack() # Inputs
frm2.pack() # Basedrive buttons
frm3.pack() # Command history
frm4.pack() # Info and checks
frm5.pack() # Disclaimer

# Setup text labels
lblCom = tk.Label(text="COM port: ", master=frm1, padx=10, pady=5)
lblBdr = tk.Label(text="Baudrate: ", master=frm1, padx=10, pady=5)
lblVerbose = tk.Label(text="Show Commands: ", master=frm1, padx=10, pady=5)
lblCadmAddr = tk.Label(text="CADM Address: ", master=frm1, padx=10, pady=5)
lblFreq = tk.Label(text="Frequency [Hz]: ",master=frm1, padx=10, pady=5)
lblRss = tk.Label(text="Step Size [%]: ",master=frm1, padx=10, pady=5)  
lblTemp = tk.Label(text="Temperature [K]: ",master=frm1, padx=10, pady=5)  
lblStage = tk.Label(text="Stage Type: ",master=frm1, padx=10, pady=5)  
lblDf = tk.Label(text="Drive Factor: ",master=frm1, padx=10, pady=5)   
lblTextBox = tk.Label(text="Command History", master=frm3, padx=10, pady=5)

lblCoeParams = tk.Label(text="COE check settings: ",master=frm1, padx=10, pady=5)
lblCoeRun = tk.Label(text="Run Number [#]: ",master=frm1, padx=10, pady=5)
lblPosId = tk.Label(text="Positioner ID [#]: ",master=frm1, padx=10, pady=5)
lblCoeId = tk.Label(text="COE ID [#]: ",master=frm1, padx=10, pady=5)
lblCoeTimerA = tk.Label(text="Timer A [s]: ",master=frm1, padx=10, pady=5)
lblCoeTimerB = tk.Label(text="Timer B [s]: ",master=frm1, padx=10, pady=5)
lblCoeFreq = tk.Label(text="Frequency [Hz]: ",master=frm1, padx=10, pady=5)

lblOemAddr = tk.Label(text="OEM Address: ", master=frm1, padx=10, pady=5)
lblOemCh = tk.Label(text="OEM Channel: ", master=frm1, padx=10, pady=5)
lblOemCal = [['']*3 for i in range(3)]
lblOemCal[0][0] = tk.Label(text="OEM Ch1 \u21E8 Gain:", master=frm1, padx=10, pady=2)
lblOemCal[0][1] = tk.Label(text="OEM Ch2 \u21E8 Gain:", master=frm1, padx=10, pady=2)
lblOemCal[0][2] = tk.Label(text="OEM Ch3 \u21E8 Gain:", master=frm1, padx=10, pady=2)
lblOemCal[1][0] = tk.Label(text="Lower Threshold:", master=frm1, padx=10, pady=2)
lblOemCal[1][1] = tk.Label(text="Lower Threshold:", master=frm1, padx=10, pady=2)
lblOemCal[1][2] = tk.Label(text="Lower Threshold:", master=frm1, padx=10, pady=2)
lblOemCal[2][0] = tk.Label(text="Upper Threshold:", master=frm1, padx=10, pady=2)
lblOemCal[2][1] = tk.Label(text="Upper Threshold:", master=frm1, padx=10, pady=2)
lblOemCal[2][2] = tk.Label(text="Upper Threshold:", master=frm1, padx=10, pady=2)

lblDisclaimer = tk.Label(text="JPE - Driven by innovation | jpe-innovations.com", master=frm5)
lblTime = tk.Label(text="", master=frm5)

# Setup error messages
errConnect = 'Cannot connect to controller!\nPlease check if CPSC1 is connected to the host via USB.\n'
   
# Declare lists
optOemCal = [[0]*3 for i in range(3)]
inpOemCal = [[0]*3 for i in range(3)]

# Set default parameter values (Spinbox and OptionMenu)
optStage = tk.StringVar(window)
optStage.set('CLA2201-COE')
optCadmAddr = tk.StringVar(window)
optCadmAddr.set('1')
optDf = tk.StringVar(window)
optDf.set('1')
optCom = tk.StringVar(window)
optCom.set('1')
optFreq = tk.StringVar(window)
optFreq.set('600')
optRss = tk.StringVar(window)
optRss.set('100')
optTemp = tk.StringVar(window)
optTemp.set('293')
optCoeRun = tk.StringVar(window)
optCoeRun.set('1')
optCoeTimerA = tk.StringVar(window)
optCoeTimerA.set('1')
optCoeTimerB = tk.StringVar(window)
optCoeTimerB.set('20')
optCoeFreq = tk.StringVar(window)
optCoeFreq.set('30')
optOemAddr = tk.StringVar(window)
optOemAddr.set('4')
optOemCh = tk.StringVar(window)
optOemCh.set('1')
optOemCal[0][0] = tk.StringVar(window) # CH1 Gain
optOemCal[0][0].set('10')
optOemCal[0][1] = tk.StringVar(window) # CH2 Gain
optOemCal[0][1].set('10')
optOemCal[0][2] = tk.StringVar(window) # CH3 Gain
optOemCal[0][2].set('10')
optOemCal[1][0] = tk.StringVar(window) # CH1 Lower Threshold
optOemCal[1][0].set('60')
optOemCal[1][1] = tk.StringVar(window) # CH2 Lower Threshold
optOemCal[1][1].set('60')
optOemCal[1][2] = tk.StringVar(window) # CH3 Lower Threshold
optOemCal[1][2].set('60')
optOemCal[2][0] = tk.StringVar(window) # CH1 Upper Threshold
optOemCal[2][0].set('180')
optOemCal[2][1] = tk.StringVar(window) # CH2 Upper Threshold
optOemCal[2][1].set('180')
optOemCal[2][2] = tk.StringVar(window) # CH3 Upper Threshold
optOemCal[2][2].set('180')
optVerbose = tk.IntVar()
optVerbose.set(False)

# Setup user inputs
inpStage = tk.OptionMenu(frm1, optStage, "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE", "CPSHR2-COE", "CPSHR3-COE")
inpCadmAddr = tk.OptionMenu(frm1, optCadmAddr, "1", "2", "3", "4", "5", "6")
inpDf = tk.OptionMenu(frm1, optDf, "0.6", "0.8", "1", "1.5", "2", "3")
inpCom = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=optCom, width=5)
inpFreq = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=optFreq, width=5)
inpRss = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=optRss, width=5)
inpTemp = tk.Spinbox(frm1, from_= 0, to = 300, textvariable=optTemp, width=5)
inpBdr = tk.Entry(width=8, master=frm1)
inpVerbose = tk.Checkbutton(frm1, text = "", variable = optVerbose, onvalue = True, offvalue = False)

inpCoeRun = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=optCoeRun, width=5)
inpPosId = tk.Entry(width=15, master=frm1)
inpCoeId = tk.Entry(width=15, master=frm1)
inpCoeTimerA = tk.Spinbox(frm1, from_= 1, to = 999, textvariable=optCoeTimerA, width=5)
inpCoeTimerB = tk.Spinbox(frm1, from_= 1, to = 999, textvariable=optCoeTimerB, width=5)
inpCoeFreq = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=optCoeFreq, width=5)

inpOemAddr = tk.OptionMenu(frm1, optOemAddr, "1", "2", "3", "4", "5", "6")
inpOemCh = tk.OptionMenu(frm1, optOemCh, "1", "2", "3")
inpOemCal[0][0] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[0][0], width=5)
inpOemCal[0][1] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[0][1], width=5)
inpOemCal[0][2] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[0][2], width=5)
inpOemCal[1][0] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[1][0], width=5)
inpOemCal[1][1] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[1][1], width=5)
inpOemCal[1][2] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[1][2], width=5)
inpOemCal[2][0] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[2][0], width=5)
inpOemCal[2][1] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[2][1], width=5)
inpOemCal[2][2] = tk.Spinbox(frm1, from_= 1, to = 255, textvariable=optOemCal[2][2], width=5)

# Set default parameter values (Entry)
inpBdr.insert(0, "115200")
inpPosId.insert(0, "12345AAA")
inpCoeId.insert(0, "12345BBB")

# Configure input box widths
inpStage.config(width=15)
inpDf.config(width=5)

# Layout of user inputs frame
lblCom.grid(row=1, column=1, sticky=tk.E)
inpCom.grid(row=1, column=2, sticky=tk.W)
lblBdr.grid(row=2, column=1, sticky=tk.E)
inpBdr.grid(row=2, column=2, sticky=tk.W)
lblVerbose.grid(row=3, column=1, sticky=tk.E)
inpVerbose.grid(row=3, column=2, sticky=tk.W)

lblCadmAddr.grid(row=1, column=3, sticky=tk.E)
inpCadmAddr.grid(row=1, column=4, sticky=tk.W)
lblOemAddr.grid(row=2, column=3, sticky=tk.E)
inpOemAddr.grid(row=2, column=4, sticky=tk.W)
lblOemCh.grid(row=3, column=3, sticky=tk.E)
inpOemCh.grid(row=3, column=4, sticky=tk.W)

lblFreq.grid(row=1, column=5, sticky=tk.E)
inpFreq.grid(row=1, column=6, sticky=tk.W)
lblRss.grid(row=2, column=5, sticky=tk.E)
inpRss.grid(row=2, column=6, sticky=tk.W)
lblTemp.grid(row=3, column=5, sticky=tk.E)
inpTemp.grid(row=3, column=6, sticky=tk.W)

lblDf.grid(row=1, column=7, sticky=tk.E)
inpDf.grid(row=1, column=8, sticky=tk.W)
lblStage.grid(row=2, column=7, sticky=tk.E)
inpStage.grid(row=2, column=8, sticky=tk.W)

lblCoeParams.grid(row=1, column=9, columnspan=2)
lblCoeRun.grid(row=2, column=9, sticky=tk.E)
inpCoeRun.grid(row=2, column=10, sticky=tk.W)
lblPosId.grid(row=3, column=9, sticky=tk.E)
inpPosId.grid(row=3, column=10, sticky=tk.W)
lblCoeId.grid(row=4, column=9, sticky=tk.E)
inpCoeId.grid(row=4, column=10, sticky=tk.W)
lblCoeTimerA.grid(row=5, column=9, sticky=tk.E)
inpCoeTimerA.grid(row=5, column=10, sticky=tk.W)
lblCoeTimerB.grid(row=6, column=9, sticky=tk.E)
inpCoeTimerB.grid(row=6, column=10, sticky=tk.W)
lblCoeFreq.grid(row=7, column=9, sticky=tk.E)
inpCoeFreq.grid(row=7, column=10, sticky=tk.W)

lblOemCal[0][0].grid(row=5, column=3, sticky=tk.E)
inpOemCal[0][0].grid(row=5, column=4, sticky=tk.W)
lblOemCal[1][0].grid(row=5, column=5, sticky=tk.E)
inpOemCal[1][0].grid(row=5, column=6, sticky=tk.W)
lblOemCal[2][0].grid(row=5, column=7, sticky=tk.E)
inpOemCal[2][0].grid(row=5, column=8, sticky=tk.W)

lblOemCal[0][1].grid(row=6, column=3, sticky=tk.E)
inpOemCal[0][1].grid(row=6, column=4, sticky=tk.W)
lblOemCal[1][1].grid(row=6, column=5, sticky=tk.E)
inpOemCal[1][1].grid(row=6, column=6, sticky=tk.W)
lblOemCal[2][1].grid(row=6, column=7, sticky=tk.E)
inpOemCal[2][1].grid(row=6, column=8, sticky=tk.W)

lblOemCal[0][2].grid(row=7, column=3, sticky=tk.E)
inpOemCal[0][2].grid(row=7, column=4, sticky=tk.W)
lblOemCal[1][2].grid(row=7, column=5, sticky=tk.E)
inpOemCal[1][2].grid(row=7, column=6, sticky=tk.W)
lblOemCal[2][2].grid(row=7, column=7, sticky=tk.E)
inpOemCal[2][2].grid(row=7, column=8, sticky=tk.W)

# Setup OEM and COE buttons 
butMls = tk.Button(text='Get OEM values', master=frm1, padx=10, pady=20)
butMss = tk.Button(text='Store OEM values', master=frm1, padx=10, pady=20)
butCoe = tk.Button(text='Start COE check', master=frm1, padx=10, pady=20)

# Layout of Basedrive and OEM buttons frame
butMls.grid(row=8, column=3, columnspan=2)
butMss.grid(row=8, column=5, columnspan=2)  
butCoe.grid(row=8, column=9, columnspan=2)  

# Layout of Command History frame
v = tk.Scrollbar(frm3, orient='vertical')
v.pack(side=tk.RIGHT, fill='y')     
lblTextBox.pack(side=tk.TOP)
txtResp = tk.Text(height = 20, master=frm3, yscrollcommand=v.set)
txtResp.tag_configure('e', foreground='red')
txtResp.tag_configure('s', foreground='blue')
txtResp.tag_configure('r', foreground='green')
txtResp.pack(side=tk.TOP)

# Setup Info and Checks buttons
butVer = tk.Button(text="Info", master=frm4, padx=10, pady=5)
butGfs = tk.Button(text="Driver Check", master=frm4, padx=10, pady=5)
butScm = tk.Button(text="Positioner Check", master=frm4, padx=10, pady=5)
butTxtRespClear = tk.Button(text="Clear command history", master=frm4, padx=10, pady=5)
butWinDevMan = tk.Button(text="Device Manager\n(Windows)", master=frm4)

# Layout of Info and Checks buttons frame
butTxtRespClear.grid(row=1, column=1, padx=10, pady=10)  
frm4.grid_columnconfigure(2, minsize=20)
butVer.grid(row=1, column=3, sticky=tk.W)
butGfs.grid(row=1, column=4, sticky=tk.W)
butScm.grid(row=1, column=5, sticky=tk.W)
frm4.grid_columnconfigure(6, minsize=20)
butWinDevMan.grid(row=1, column=7, sticky=tk.W)

# Layout of Disclaimer frame
lblDisclaimer.grid(row=1, column=1)
lblTime.grid(row=1, column=2)

# Functions
def butVer_handle_click(event):
   cmdVer = '/VER'
   cmdFiv = 'FIV '
   txtResp.insert('end', ('==> Get firmware version information\n'))
   try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:   
           if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdVer + '\n'), 's') 
           response = usbVcp.WriteRead(cmdVer, 1)
           txtResp.insert('end', ('<== CPSC1 version: ' + response + '\n'), 'r')
           for x in range(6):
               if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdFiv + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdFiv + str(x+1), 1)
               txtResp.insert('end', ('<== Slot ' + str(x+1) + ': ' + response + '\n'), 'r')            
           txtResp.see("end")
   except IOError:
       txtResp.insert('end', errConnect, 'e')     

def butGfs_handle_click(event):
    cmdGfs = 'GFS '
    txtResp.insert('end', ('==> Get CADM2 failsafe state\n'))
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:    
           for x in range(6):
               if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdGfs + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdGfs + str(x+1), 1)
               txtResp.insert('end', ('<== Slot ' + str(x+1) + ': ' + response + '\n'), 'r')
           txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')      

def butScm_handle_click(event):
    butScmThrd=thrd.Thread(target=butScm_handle_thread)
    butScmThrd.start()

def butScm_handle_thread():  
    cmdScm = 'SCM '
    txtResp.insert('end', ('==> Check piezo in positioners (will take some time) ... \n'))
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:    
           for x in range(6):
               if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdScm + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdScm + str(x+1), 1)
               txtResp.insert('end', ('<== Positioner on slot ' + str(x+1) + ': ' + response + '[nF]\n'), 'r')
           txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')     

def butMls_handle_click(event):
    butMlsThrd=thrd.Thread(target=butMls_handle_thread)
    butMlsThrd.start()
    
def butMls_handle_thread():  
    cmdMls = 'MLS ' + str(optOemAddr.get()) + ' '
    txtResp.insert('end', ('==> Get current OEM calibration values.\n'))
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:   
            for x in range(3):
                if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdMls + str(x+1) + '\n'), 's')
                response = usbVcp.WriteRead(cmdMls + str(x+1), 1)
                responseOem = response.split(",")  
                txtResp.insert('end', ('<== Ch' + str(x+1) + ': Gain: ' + responseOem[0] + '  Lower Threshold: ' + responseOem[2] + '  Upper Threshold: ' + responseOem[1] + '\n'), 'r')            
            txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')     

def butMss_handle_click(event):
    butMssThrd=thrd.Thread(target=butMss_handle_thread)
    butMssThrd.start()
    
def butMss_handle_thread():   
    cmdDsg = 'DSG ' + str(optOemAddr.get()) + ' '
    cmdDsl = 'DSL ' + str(optOemAddr.get()) + ' '
    cmdDsh = 'DSH ' + str(optOemAddr.get()) + ' '
    cmdMss = 'MSS ' + str(optOemAddr.get()) + ' '
    txtResp.insert('end', ('==> Store selected OEM calibration values.\n'))
      
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:   
            for x in range(3):
                if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdDsg + str(x+1) + '\n'), 's')
                response = usbVcp.WriteRead(cmdDsg + str(x+1) + ' ' + str(inpOemCal[0][x].get()), 1)
                txtResp.insert('end', ('<== ' + response + '\n'), 'r') 
                if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdDsl + str(x+1) + '\n'), 's')
                response = usbVcp.WriteRead(cmdDsl + str(x+1) + ' ' + str(inpOemCal[1][x].get()), 1)
                txtResp.insert('end', ('<== ' + response + '\n'), 'r') 
                if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdDsh + str(x+1) + '\n'), 's')
                response = usbVcp.WriteRead(cmdDsh + str(x+1) + ' ' + str(inpOemCal[2][x].get()), 1)
                txtResp.insert('end', ('<== ' + response + '\n'), 'r')            
            response = usbVcp.WriteRead(cmdMss, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
        txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')  

def butCoe_handle_click(event): 
    butCoeThrd=thrd.Thread(target=butCoe_handle_thread)
    butCoeThrd.start()
    
def butCoe_handle_thread():    
    startTime = 0
    passedTime = 0
    responseTime = []
    responseDgv = []
    responseCgv = []
    responseOem = []
    pollDelay = 0.5
    
    cmdMls = 'MLS ' + str(optOemAddr.get()) + ' ' + str(optOemCh.get())
    cmdMvA = 'MOV ' + str(optCadmAddr.get()) + ' 0 ' + str(optFreq.get()) + ' ' + str(optRss.get()) + ' 0 ' + str(optTemp.get()) + ' ' + str(optStage.get()) + ' ' + str(optDf.get())
    cmdMvB = 'MOV ' + str(optCadmAddr.get()) + ' 1 ' + str(optCoeFreq.get()) + ' ' + str(optRss.get()) + ' 0 ' + str(optTemp.get()) + ' ' + str(optStage.get()) + ' ' + str(optDf.get())
    cmdStp = 'STP ' + str(optCadmAddr.get())
    cmdCsz = 'CSZ ' + str(optOemAddr.get()) + ' ' + str(optOemCh.get())
    cmdDgv = 'DGV ' + str(optOemAddr.get()) + ' ' + str(optOemCh.get())
    cmdCgv = 'CGV ' + str(optOemAddr.get()) + ' ' + str(optOemCh.get())

    txtResp.insert('end', ('==> Start COE check. Please be patient!\n'))
       
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:   
            txtResp.insert('end', ('==> Get current OEM calibration values.\n'))  
            if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdMls + '\n'), 's')
            response = usbVcp.WriteRead(cmdMls, 1)
            responseOem = response.split(",")          
            txtResp.insert('end', ('<== Ch' + str(optOemCh.get()) + ': Gain: ' + responseOem[0] + '  Lower Threshold: ' + responseOem[2] + '  Upper Threshold: ' + responseOem[1] + '\n'), 'r')            

            txtResp.insert('end', ('==> Start moving in DIR=0 for ' + optCoeTimerA.get() + '[s]\n'))
            if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdMvA + '\n'), 's')
            startTime = time.time()
            response = usbVcp.WriteRead(cmdMvA, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
            
            while not (passedTime > int(optCoeTimerA.get())):
                passedTime = time.time() - startTime
                txtResp.insert('end', ('<== ET: ' + str(round((passedTime),1)) + '\n'), 'r')
                txtResp.see("end")
                time.sleep(pollDelay)

            txtResp.insert('end', ('==> Stop moving\n'))
            if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
            response = usbVcp.WriteRead(cmdStp, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
    
            txtResp.insert('end', ('==> Reset counter to zero\n'))
            if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdCsz + '\n'), 's')
            response = usbVcp.WriteRead(cmdCsz, 1) # Reset counter
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
    
            txtResp.insert('end', ('==> Start moving in DIR=1 for ' + optCoeTimerB.get() + '[s]\n'))
            if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdMvB + '\n'), 's')
            passedTime = 0
            startTime = time.time()
            response = usbVcp.WriteRead(cmdMvB, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
            
            while not (passedTime > int(optCoeTimerB.get())):        
                responseDgvTemp = usbVcp.WriteRead(cmdDgv, 1)
                responseDgv.append(int(responseDgvTemp))
                responseCgvTemp = usbVcp.WriteRead(cmdCgv, 1)
                responseCgv.append(int(responseCgvTemp))       
                passedTime = time.time() - startTime
                responseTime.append(passedTime)
                txtResp.insert('end', ('<== ET: ' + str(round((passedTime),1)) + ', CGV: ' + responseCgvTemp + ', DGV: ' + responseDgvTemp + '\n'), 'r')
                txtResp.see("end")
   
            txtResp.insert('end', ('==> Stop moving\n'))
            if (optVerbose.get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
            response = usbVcp.WriteRead(cmdStp, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')   
            txtResp.see("end")
               
    except IOError:
        txtResp.insert('end', errConnect, 'e')       
           
    txtResp.insert('end', ('==> Create plot and store data to file ... \n'))
    fig, (ax1) = plt.subplots(1)
    fig.suptitle(str(optStage.get()) + ' #' + str(inpPosId.get()) + ' | COE #' + str(inpCoeId.get()) + ' | Run #' + str(optCoeRun.get()) + 
                 '\nTest Param: [FREQ]:' + str(optCoeFreq.get()) + ' [RSS]:' + str(optRss.get()) + ' [TEMP]:' + str(optTemp.get()) + ' [DF]:' + str(optDf.get()) + 
                 ' | OEM [GAIN]:' + responseOem[0] + ' [UT]:' + responseOem[1] + ' [LT]:' + responseOem[2] +
                 '\n(measured around mid-stroke of COE)', fontsize=8)
    ax1.set_xlabel('Time [s]')
    ax1.set_ylabel('COE POS [counts]', color='r')
    ax1.plot(responseTime, responseCgv, color='r', linewidth=0.5)
    ax1.tick_params(axis='y', labelcolor='r')
    ax1.grid(linestyle = '--', linewidth = 0.2)
    ax2 = ax1.twinx()
    ax2.set_ylabel('COE RAW value', color='b')
    ax2.plot(responseTime, responseDgv, color='b', linewidth=0.5)
    ax2.tick_params(axis='y', labelcolor='b')
    ax2.axhline(y = int(responseOem[1]), color = 'tab:gray', linestyle = 'dashed')
    ax2.axhline(y = int(responseOem[2]), color = 'tab:gray', linestyle = 'dashed')
    fig.tight_layout()
     
    plt.savefig(str(optStage.get()) + '_' + str(inpPosId.get()) + '_' + str(inpCoeId.get()) + '-' + str(optCoeRun.get()) + '.png', dpi=300)
    plt.show()
    
    #plt.close()
     
    TimeData = list(zip(responseTime, responseCgv, responseDgv))
    
    with open(str(optStage.get()) + '_' + str(inpPosId.get()) + '_' + str(inpCoeId.get()) + '-' + str(optCoeRun.get()) +'.csv', 'w', newline="") as x:
        csv.writer(x, delimiter=";", dialect='excel').writerows(TimeData)
    x.close()        

    txtResp.insert('end', ('==> COE Count: ' + str(max(responseCgv)) + '\n'))
    txtResp.insert('end', ('==> Duration: ' + str(round((passedTime),1)) + '\n'))
    txtResp.insert('end', ('==> ** COE check completed **\n'))   
    txtResp.see("end")
      
def txtResp_clear_click(event):
    txtResp.delete("1.0", "end")

def butWinDevMan_clear_click(event):
    sp.call("control /name Microsoft.DeviceManager")

# Bind button press (left mouse click) events to functions
butVer.bind("<Button-1>", butVer_handle_click)
butGfs.bind("<Button-1>", butGfs_handle_click)
butScm.bind("<Button-1>", butScm_handle_click)
butMls.bind("<Button-1>", butMls_handle_click)
butMss.bind("<Button-1>", butMss_handle_click)
butCoe.bind("<Button-1>", butCoe_handle_click)
butTxtRespClear.bind("<Button-1>", txtResp_clear_click)
butWinDevMan.bind("<Button-1>", butWinDevMan_clear_click)

# Display some initial tips and hints
txtResp.insert('end', ('==> Note: check CADM Address and OEM Address/Channel settings and start with requesting current OEM settings ("Get OEM values").\n'))
txtResp.insert('end', ('==> Run "Start COE check" to see if these values are okay for the connected stage type (a graph will be created and stored in the script folder).\n'))
txtResp.insert('end', ('==> If necessary update Gain and Threshold values and store settings ("Store OEM values") and run "Start COE check" again.\n'))

# Main loop (loop until window is closed)
window.mainloop()
