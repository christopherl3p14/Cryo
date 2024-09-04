###############################################################################
# File name:      CPSC1_BaseDrive_GUI-vX.y.py
# Creation:       2022-12-07
# Last Update:    2024-03-22
# Author:         JPE
# Python version: 3.9, requires pyserial and tkinter
# Description:    MVP GUI controlling CPSC1 "Basedrive" mode of operation
#                 for up to 6 positioners (6x CADM)
# Disclaimer:     This program is provided 'As Is' without any express or 
#                 implied warranty of any kind.
###############################################################################

verNumber = 'v0.3'

# Defaults for drop down lists
lstStages = ['CBS5', 'CBS5-HF', 'CBS10', 'CLA2201', 'CLA2601', 'CLA2602', 'CLA2603', 'CSR1', 'CSR2', 'CRM1', 'CLD2', 'CPSHR1-S', 'CPSHR2', 'CPSHR3', 'CTTPS1/2', 'CTTPS1']
lstDf = ['0.5','0.6','0.7','0.8','0.9','1.0','1.1','1.2','1.3','1.4','1.5','2.0','2.5','3.0']
lstCh = ['1','2','3','4','5','6']

# 3rd party imports
import time
import tkinter as tk
import subprocess as sp
import threading as thrd

# JPE imports
from CpscInterfaces import CpscSerialInterface as CpscSerial

# Create GUI window
window = tk.Tk()

# Setup icon global window parameters
p1 = tk.PhotoImage(file = 'jpe.png')
window.iconphoto(False, p1)
window.title(f'CPSC1 BaseDrive for up to 6 positioners ({verNumber})') 
window.option_add( "*font", "Helvetica 11" )

# Setup frames within window
frm1 = tk.Frame(width=500, padx=10, pady=10)
frm2 = tk.Frame(width=500, padx=10, pady=10)
frm3 = tk.Frame(width=500, height=400, padx=10, pady=10) 
frm4 = tk.Frame(width=500, padx=10, pady=10) 
frm5 = tk.Frame(width=500, padx=10, pady=10) 
frm1.pack() # Inputs
frm2.pack() # Basedrive buttons
frm3.pack() # Command history
frm4.pack() # Info and checks
frm5.pack() # Disclaimer

# Setup text labels
lblCom = tk.Label(text="COM port: ", master=frm1, padx=10, pady=10)
lblBdr = tk.Label(text="Baudrate: ", master=frm1, padx=10, pady=10)
lblVerbose = tk.Label(text="Show Commands: ", master=frm1, padx=10, pady=5)
lblAddr = tk.Label(text="Address: ", master=frm1, padx=10, pady=10)
lblFreq = tk.Label(text="Frequency [Hz]: ",master=frm1, padx=10, pady=10)
lblRss = tk.Label(text="Step Size [%]: ",master=frm1, padx=10, pady=10)  
lblStep = tk.Label(text="[No.] of Steps: ",master=frm1, padx=10, pady=10)  
lblTemp = tk.Label(text="Temperature [K]: ",master=frm1, padx=10, pady=10)  
lblStage = tk.Label(text="Stage Type: ",master=frm1, padx=10, pady=10)  
lblDf = tk.Label(text="Drive Factor: ",master=frm1, padx=10, pady=10)   
lblTextBox = tk.Label(text="Command History", master=frm3, padx=10, pady=10)
lblDisclaimer = tk.Label(text="JPE - Driven by innovation | jpe-innovations.com", master=frm5)
lblTime = tk.Label(text="", master=frm5)

# Setup error messages
errConnect = 'Cannot connect to controller!\nPlease check if CPSC1 is connected to the host via USB and the correct COM port number and Baudrate has been set.\n'
   
# Set default parameter values (Spinbox and OptionMenu)
optStage = tk.StringVar(window)
optStage.set('CLA2201')
optAddr = tk.StringVar(window)
optAddr.set('1')
optDf = tk.StringVar(window)
optDf.set('1.0')
optCom = tk.StringVar(window)
optCom.set('1')
optFreq = tk.StringVar(window)
optFreq.set('600')
optRss = tk.StringVar(window)
optRss.set('100')
optTemp = tk.StringVar(window)
optTemp.set('293')
optVerbose = tk.IntVar()
optVerbose.set(True)

# Setup user inputs
inpStage = tk.OptionMenu(frm1, optStage, *lstStages)
inpAddr = tk.OptionMenu(frm1, optAddr, *lstCh)
inpDf = tk.OptionMenu(frm1, optDf, *lstDf)
inpCom = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=optCom, width=5)
inpFreq = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=optFreq, width=5)
inpRss = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=optRss, width=5)
inpTemp = tk.Spinbox(frm1, from_= 0, to = 300, textvariable=optTemp, width=5)
inpStep = tk.Entry(width=6, master=frm1)
inpBdr = tk.Entry(width=8, master=frm1)
inpVerbose = tk.Checkbutton(frm1, text = "", variable = optVerbose, onvalue = True, offvalue = False)

# Set default parameter values (Entry)
inpStep.insert(0, "0")
inpBdr.insert(0, "115200")         

# Configure input box widths
inpStage.config(width=8)
inpDf.config(width=8)

# Layout of user inputs frame
lblCom.grid(row=1, column=1, sticky=tk.E)
inpCom.grid(row=1, column=2, sticky=tk.W)
lblBdr.grid(row=2, column=1, sticky=tk.E)
inpBdr.grid(row=2, column=2, sticky=tk.W)
lblVerbose.grid(row=3, column=1, sticky=tk.E)
inpVerbose.grid(row=3, column=2, sticky=tk.W)

lblAddr.grid(row=1, column=3, sticky=tk.E)
inpAddr.grid(row=1, column=4, sticky=tk.W)

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
lblStep.grid(row=3, column=7, sticky=tk.E)
inpStep.grid(row=3, column=8, sticky=tk.W)

# Setup Basedrive buttons   
butStp = tk.Button(text="STOP [End]", master=frm2, padx=10, pady=20)
butDir0 = tk.Button(text="DIR=0 [PgUp]", master=frm2, padx=10, pady=20)
butDir1 = tk.Button(text="DIR=1 [PgDw]", master=frm2, padx=10, pady=20)

# Layout of Basedrive and Manual Rise Time buttons frame
butDir0.grid(row=1, column=1, sticky=tk.E)
butStp.grid(row=1, column=2, sticky=tk.E)
butDir1.grid(row=1, column=3, sticky=tk.E)
frm2.grid_columnconfigure(4, minsize=50)
     
# Layout of Command History frame
v = tk.Scrollbar(frm3, orient='vertical')
v.pack(side=tk.RIGHT, fill='y')     
lblTextBox.pack(side=tk.TOP)
txtRespFont = ("Courier New", 12)
txtResp = tk.Text(height=15, width=60, master=frm3, yscrollcommand=v.set)
txtResp.tag_configure('e', foreground='red')
txtResp.tag_configure('s', foreground='blue')
txtResp.tag_configure('r', foreground='green')
txtResp.configure(font = txtRespFont)
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
   butVerThrd=thrd.Thread(target=butVer_handle_thread)
   butVerThrd.start()
    
def butVer_handle_thread():  
   cmdVer = '/VER'
   cmdFiv = 'FIV '
   txtResp.insert('end', ('--> Get firmware version information\n'))
   if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdVer + '\n'), 's') 
   try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:   
           response = usbVcp.WriteRead(cmdVer, 1)
           txtResp.insert('end', ('<-- CPSC1 version: ' + response + '\n'), 'r')
           for x in range(6):
               if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdFiv + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdFiv + str(x+1), 1)
               txtResp.insert('end', ('<-- Slot ' + str(x+1) + ': ' + response + '\n'), 'r')            
               txtResp.see("end")
   except IOError:
       txtResp.insert('end', errConnect, 'e')      

def butGfs_handle_click(event):
    butGfsThrd=thrd.Thread(target=butGfs_handle_thread)
    butGfsThrd.start()
     
def butGfs_handle_thread():  
    cmdGfs = 'GFS '
    txtResp.insert('end', ('--> Get CADM2 failsafe state\n'))
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:    
           for x in range(6):
               if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdGfs + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdGfs + str(x+1), 1)
               txtResp.insert('end', ('<-- Slot ' + str(x+1) + ': ' + response + '\n'), 'r')
               txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')        

def butScm_handle_click(event):
    butScmThrd=thrd.Thread(target=butScm_handle_thread)
    butScmThrd.start()

def butScm_handle_thread():  
    cmdScm = 'SCM '
    txtResp.insert('end', ('--> Check piezo in positioners (will take some time) ... \n'))
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:    
           for x in range(6):
               if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdScm + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdScm + str(x+1), 1)
               txtResp.insert('end', ('<-- Positioner on slot ' + str(x+1) + ': ' + response + '[nF]\n'), 'r')
               txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')     
        
def butStp_handle_click(event):
    butStpThrd=thrd.Thread(target=butStp_handle_thread)
    butStpThrd.start()
     
def butStp_handle_thread(): 
    cmdStp = 'STP ' + str(optAddr.get())
    txtResp.insert('end', ('--> Stop movement\n'))
    if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdStp + '\n'), 's')
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:          
           response = usbVcp.WriteRead(cmdStp, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')       

def butDir0_handle_click(event):
    butDir0Thrd=thrd.Thread(target=butDir0_handle_thread)
    butDir0Thrd.start()
     
def butDir0_handle_thread():     
    dirMov = '0'
    cmdMov = 'MOV ' + str(optAddr.get()) + ' ' + dirMov + ' ' + str(optFreq.get()) + ' ' + str(optRss.get()) + ' ' + str(inpStep.get()) + ' ' + str(optTemp.get()) + ' ' + str(optStage.get()) + ' ' + str(optDf.get())
    txtResp.insert('end', ('--> Start movement in DIR=0 direction\n'))
    if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdMov + '\n'), 's')
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:         
           response = usbVcp.WriteRead(cmdMov, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
       txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')    

def butDir1_handle_click(event):
    butDir1Thrd=thrd.Thread(target=butDir1_handle_thread)
    butDir1Thrd.start()
     
def butDir1_handle_thread():  
    dirMov = '1'
    cmdMov = 'MOV ' + str(optAddr.get()) + ' ' + dirMov + ' ' + str(optFreq.get()) + ' ' + str(optRss.get()) + ' ' + str(inpStep.get()) + ' ' + str(optTemp.get()) + ' ' + str(optStage.get()) + ' ' + str(optDf.get())
    txtResp.insert('end', ('--> Start movement in DIR=1 direction\n'))
    if (optVerbose.get()): txtResp.insert('end', ('--> ' + cmdMov + '\n'), 's')
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optCom.get())), str(inpBdr.get())) as usbVcp:    
           response = usbVcp.WriteRead(cmdMov, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
       txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')    

def txtResp_clear_click(event):
    txtResp.delete("1.0", "end")

def butWinDevMan_clear_click(event):
    sp.call("control /name Microsoft.DeviceManager")

def doNothing(event):
    time.sleep(1)

# Bind button press (left mouse click) events to functions
butVer.bind("<Button-1>", butVer_handle_click)
butGfs.bind("<Button-1>", butGfs_handle_click)
butScm.bind("<Button-1>", butScm_handle_click)
butStp.bind("<Button-1>", butStp_handle_click)
butDir1.bind("<Button-1>", butDir1_handle_click)
butDir0.bind("<Button-1>", butDir0_handle_click)
butTxtRespClear.bind("<Button-1>", txtResp_clear_click)
butWinDevMan.bind("<Button-1>", butWinDevMan_clear_click)

butVer.bind("<Double-Button-1>", doNothing)
butGfs.bind("<Double-Button-1>", doNothing)
butScm.bind("<Double-Button-1>", doNothing)
butStp.bind("<Double-Button-1>", doNothing)
butDir0.bind("<Double-Button-1>", doNothing)
butDir1.bind("<Double-Button-1>", doNothing)

# Bine key press events to functions  
window.bind("<End>", butStp_handle_click)  
window.bind("<Next>", butDir1_handle_click)
window.bind("<Prior>", butDir0_handle_click)      
 
# Main loop (loop until window is closed)
window.mainloop()
