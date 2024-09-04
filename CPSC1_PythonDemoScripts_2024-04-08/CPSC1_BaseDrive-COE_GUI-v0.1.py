###############################################################################
# File name:      CPSC1_BaseDrive-COE_GUI_vX.y.py
# Creation:       2023-01-31
# Last Update:    2023-01-31
# Author:         JPE
# Python version: 3.9, requires pyserial and tkinter
# Description:    MVP GUI controlling CPSC1 "Basedrive" mode of operation
#                 for up to 3 positioners with COE (3x CADM + 1x OEM)
# Disclaimer:     This program is provided 'As Is' without any express or 
#                 implied warranty of any kind.
###############################################################################

verNumber = 'v0.1'

# 3rd party imports
import time
import tkinter as tk
import subprocess as sp
import threading as thrd
from queue import Queue

# JPE imports
from CpscInterfaces import CpscSerialInterface as CpscSerial

# Create GUI window
window = tk.Tk()

# Setup icon global window parameters
p1 = tk.PhotoImage(file = 'jpe.png')
window.iconphoto(False, p1)
window.title(f'CPSC1 BaseDrive for 3 positioners with COE ({verNumber})') 
window.option_add( "*font", "Helvetica 11" )

# Setup frames within window
frm1 = tk.Frame(width=500, padx=40, pady=10)
frm2 = tk.Frame(width=500, padx=40, pady=10)
frm3 = tk.Frame(width=500, padx=40, pady=10) 
frm4 = tk.Frame(width=500, padx=40, pady=10) 
frm5 = tk.Frame(width=500, padx=40, pady=10) 
frm6 = tk.Frame(width=500, padx=40, pady=10) 
frm1.pack() # General Inputs
frm2.pack() # RSM Inputs
frm3.pack() # Basedrive buttons
frm4.pack() # Command history
frm5.pack() # Info and checks
frm6.pack() # Disclaimer

# # Setup param list defaults
parList = [[0]*8 for i in range(4)]
parList[0][0] = tk.StringVar(window) # Channel 1, CADM address
parList[0][0].set('1')
#parList[1][0] = tk.StringVar(window) # Channel 2, CADM address
#parList[1][0].set('2')
#parList[2][0] = tk.StringVar(window) # Channel 3, CADM address
#parList[2][0].set('3')
parList[0][1] = tk.StringVar(window) # Channel 1, Direction of movement
parList[0][1].set('1')
#parList[1][1] = tk.StringVar(window) # Channel 2, Direction of movement
#parList[1][1].set('1')
#parList[2][1] = tk.StringVar(window) # Channel 3, Direction of movement
#parList[2][1].set('1')
parList[0][2] = tk.StringVar(window) # Channel 1, Frequency of movement
parList[0][2].set('600')
#parList[1][2] = tk.StringVar(window) # Channel 2, Frequency of movement
#parList[1][2].set('600')
#parList[2][2] = tk.StringVar(window) # Channel 3, Frequency of movement
#parList[2][2].set('600')
parList[0][3] = tk.StringVar(window) # Channel 1, Relative Step Size of movement
parList[0][3].set('100')
#parList[1][3] = tk.StringVar(window) # Channel 2, Relative Step Size of movement
#parList[1][3].set('100')
#parList[2][3] = tk.StringVar(window) # Channel 3, Relative Step Size of movement
#parList[2][3].set('100')
parList[0][4] = tk.StringVar(window) # Channel 1, Number of steps of movement
parList[0][4].set('0')
#parList[1][4] = tk.StringVar(window) # Channel 2, Number of steps of movement
#parList[1][4].set('0')
#parList[2][4] = tk.StringVar(window) # Channel 3, Number of steps of movement
#parList[2][4].set('0')
parList[0][5] = tk.StringVar(window) # Channel 1, Connected stage type (movement)
parList[0][5].set('CLA2201-COE')
#parList[1][5] = tk.StringVar(window) # Channel 2, Connected stage type (movement)
#parList[1][5].set('CLA2201-COE')
#parList[2][5] = tk.StringVar(window) # Channel 3, Connected stage type (movement)
#parList[2][5].set('CLA2201-COE')
parList[0][6] = tk.StringVar(window) # Channel 1, Drive Factor for movement
parList[0][6].set('1')
#parList[1][6] = tk.StringVar(window) # Channel 2, Drive Factor for movement
#parList[1][6].set('1')
#parList[2][6] = tk.StringVar(window) # Channel 3, Drive Factor for movement
#parList[2][6].set('1')
parList[3][0] = tk.StringVar(window) # COM Port
parList[3][0].set('1')
parList[3][1] = tk.StringVar(window) # Environment temperature
parList[3][1].set('293')
parList[3][2] = tk.StringVar(window) # OEM Address
parList[3][2].set('2')
parList[3][3] = tk.IntVar()          # Show commands toggle
parList[3][3].set(False)
parList[3][5] = tk.IntVar()          # OEM toggle
parList[3][5].set(False)
parList[3][6] = tk.IntVar()          # Command lock
parList[3][6].set(False)


# Setup text DRIVE labels
lblList = [0]*22
lblList[0] = tk.Label(text="COM port: ", master=frm1)
lblList[1] = tk.Label(text="Baudrate: ", master=frm1)
lblList[2] = tk.Label(text="Show Commands: ", master=frm1)
lblList[3] = tk.Label(text="OEM Address: ", master=frm1)
lblList[4] = tk.Label(text="CADM Addr: ", master=frm1)
lblList[5] = tk.Label(text="Direction: ",master=frm1)
lblList[6] = tk.Label(text="Freq [Hz]: ",master=frm1)
lblList[7] = tk.Label(text="Rss [%]: ",master=frm1)  
lblList[8] = tk.Label(text="[#] Steps: ",master=frm1) 
lblList[9] = tk.Label(text="Temperature [K]: ",master=frm1) 
lblList[10] = tk.Label(text="Stage Type: ",master=frm1) 
lblList[11] = tk.Label(text="Drive Factor: ",master=frm1) 
lblList[12] = tk.Label(text="CH1: ",master=frm1) 
#lblList[13] = tk.Label(text="CH2: ",master=frm1) 
#lblList[14] = tk.Label(text="CH3: ",master=frm1) 
lblList[16] = tk.Label(text="Command History", master=frm4)
lblList[17] = tk.Label(text="JPE - Driven by innovation | jpe-innovations.com", master=frm6)

# Setup COE/OEM data list
oemList = [[0]*4 for i in range(4)]
oemList[0][0] = tk.Label(text="OEM Enable: ",master=frm2) 
oemList[1][0] = tk.Label(text="CH1 counter: ",master=frm2) 
#oemList[2][0] = tk.Label(text="CH2 counter: ",master=frm2) 
#oemList[3][0] = tk.Label(text="CH3 counter: ",master=frm2) 
oemList[1][1] = tk.Label(text="0", width = 6, master=frm2) 
oemList[2][1] = tk.Label(text="0", width = 6, master=frm2) 
oemList[3][1] = tk.Label(text="0", width = 6, master=frm2) 

# Setup error messages
errConnect = 'Cannot connect to controller!\nPlease check if CPSC1 is connected to the host via USB and the correct COM port number and Baudrate has been set.\n'

# Setup user inputs
inpList = [[0]*8 for i in range(4)]
inpList[0][0] = tk.OptionMenu(frm1, parList[0][0], "1", "2", "3", "4", "5", "6")
inpList[1][0] = tk.OptionMenu(frm1, parList[1][0], "1", "2", "3", "4", "5", "6")
inpList[2][0] = tk.OptionMenu(frm1, parList[2][0], "1", "2", "3", "4", "5", "6")
inpList[0][1] = tk.OptionMenu(frm1, parList[0][1], "0", "1")
inpList[1][1] = tk.OptionMenu(frm1, parList[1][1], "0", "1")
inpList[2][1] = tk.OptionMenu(frm1, parList[2][1], "0", "1")
inpList[0][2] = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=parList[0][2], width=5)
inpList[1][2] = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=parList[1][2], width=5)
inpList[2][2] = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=parList[2][2], width=5)
inpList[0][3] = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=parList[0][3], width=5)
inpList[1][3] = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=parList[1][3], width=5)
inpList[2][3] = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=parList[2][3], width=5)
inpList[0][4] = tk.Entry(width=6, master=frm1)
inpList[1][4] = tk.Entry(width=6, master=frm1)
inpList[2][4] = tk.Entry(width=6, master=frm1)
inpList[0][5] = tk.OptionMenu(frm1, parList[0][5], "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE", "CPSHR2-COE", "CPSHR3-COE")
inpList[1][5] = tk.OptionMenu(frm1, parList[1][5], "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE", "CPSHR2-COE", "CPSHR3-COE")
inpList[2][5] = tk.OptionMenu(frm1, parList[2][5], "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE", "CPSHR2-COE", "CPSHR3-COE")
inpList[0][6] = tk.OptionMenu(frm1, parList[0][6], "0.6", "0.8", "1", "1.5", "2", "3")
inpList[1][6] = tk.OptionMenu(frm1, parList[1][6], "0.6", "0.8", "1", "1.5", "2", "3")
inpList[2][6] = tk.OptionMenu(frm1, parList[2][6], "0.6", "0.8", "1", "1.5", "2", "3")
inpList[3][0] = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=parList[3][0], width=5)
inpList[3][1] = tk.Spinbox(frm1, from_= 0, to = 300, textvariable=parList[3][1], width=5)
inpList[3][2] = tk.OptionMenu(frm1, parList[3][2], "1", "2", "3", "4", "5", "6")
inpList[3][3] = tk.Checkbutton(frm1, text = "", variable = parList[3][3], onvalue = True, offvalue = False)
inpList[3][4] = tk.Entry(width=8, master=frm1)
inpList[3][5] = tk.Checkbutton(frm2, text = "", variable = parList[3][5], onvalue = True, offvalue = False)

# Set default parameter values for (Entry) widgets
inpList[0][4].insert(0, "0") # CH1 No of Steps
inpList[1][4].insert(0, "0") # CH1 No of Steps
inpList[2][4].insert(0, "0") # CH1 No of Steps
inpList[3][4].insert(0, "115200") # BaudRate         

# Setup Move and Stop buttons   
butMov1 = tk.Button(text="Move [Ins]", master=frm1, padx=20)
butMov2 = tk.Button(text="Move [Hom]", master=frm1, padx=20)
butMov3 = tk.Button(text="Move [PgU]", master=frm1, padx=20)
butStp1 = tk.Button(text="Stop [Del]", master=frm1, padx=20)
butStp2 = tk.Button(text="Stop [End]", master=frm1, padx=20)
butStp3 = tk.Button(text="Stop [PgD]", master=frm1, padx=20)

# Configure input box widths
inpList[0][0].config(width=5)
inpList[1][0].config(width=5)
inpList[2][0].config(width=5)                       
inpList[0][1].config(width=5)
inpList[1][1].config(width=5)
inpList[2][1].config(width=5)                       
inpList[0][5].config(width=12)
inpList[1][5].config(width=12)
inpList[2][5].config(width=12)
inpList[0][6].config(width=5)
inpList[1][6].config(width=5)
inpList[2][6].config(width=5)                       

# Layout of user inputs frame
lblList[0].grid(row=1, column=1) # Label COM port
inpList[3][0].grid(row=1, column=2)
lblList[1].grid(row=1, column=3) # Label Baudrate
inpList[3][4].grid(row=1, column=4)
lblList[2].grid(row=1, column=5) # Label Verbose
inpList[3][3].grid(row=1, column=6)
lblList[9].grid(row=1, column=7) # Label Temperature
inpList[3][1].grid(row=1, column=8)
lblList[3].grid(row=1, column=9) # Label RSM address
inpList[3][2].grid(row=1, column=10)
                   
lblList[4].grid(row=2, column=2) # Label CADM address
lblList[5].grid(row=2, column=3) # Label Direction
lblList[6].grid(row=2, column=4) # Label Frequency
lblList[7].grid(row=2, column=5) # Label Step Size
lblList[8].grid(row=2, column=6) # Label No of Stps
lblList[10].grid(row=2, column=7) # Label Stage Type
lblList[11].grid(row=2, column=8) # Label Drive Factor

lblList[12].grid(row=3, column=1) # Channel 1 label
inpList[0][0].grid(row=3, column=2)
inpList[0][1].grid(row=3, column=3)
inpList[0][2].grid(row=3, column=4)
inpList[0][3].grid(row=3, column=5)
inpList[0][4].grid(row=3, column=6)
inpList[0][5].grid(row=3, column=7)
inpList[0][6].grid(row=3, column=8)
butMov1.grid(row=3, column=9)
butStp1.grid(row=3, column=10)

#lblList[13].grid(row=4, column=1) # Channel 2 label
#inpList[1][0].grid(row=4, column=2)
#inpList[1][1].grid(row=4, column=3)
#inpList[1][2].grid(row=4, column=4)
#inpList[1][3].grid(row=4, column=5)
#inpList[1][4].grid(row=4, column=6)
#inpList[1][5].grid(row=4, column=7)
#inpList[1][6].grid(row=4, column=8)
#butMov2.grid(row=4, column=9)
#butStp2.grid(row=4, column=10)

#lblList[14].grid(row=5, column=1) # Channel 3 label
#inpList[2][0].grid(row=5, column=2)
#inpList[2][1].grid(row=5, column=3)
#inpList[2][2].grid(row=5, column=4)
#inpList[2][3].grid(row=5, column=5)
#inpList[2][4].grid(row=5, column=6)
#inpList[2][5].grid(row=5, column=7)
#inpList[2][6].grid(row=5, column=8)
#butMov3.grid(row=5, column=9)
#butStp3.grid(row=5, column=10)

oemList[0][0].grid(row=0, column=0) # OEM Readout Enable label
inpList[3][5].grid(row=0, column=1) # OEM Readout Enable
oemList[1][0].grid(row=0, column=2) # CH1 label
#oemList[2][0].grid(row=1, column=2) # CH2 label
#oemList[3][0].grid(row=2, column=2) # CH3 label
oemList[1][1].grid(row=0, column=3) # CH1 CNT val
#oemList[2][1].grid(row=1, column=3) # CH2 CNT val
#oemList[3][1].grid(row=2, column=3) # CH3 CNT val

# Setup OEM reset buttons
butCsz1 = tk.Button(text="Reset", master=frm2)
#butCsz2 = tk.Button(text="Reset", master=frm2)
#butCsz3 = tk.Button(text="Reset", master=frm2)
butCsz1.grid(row=0, column=4)
#butCsz2.grid(row=1, column=4)
#butCsz3.grid(row=2, column=4)
     
# Layout of Command History frame
v = tk.Scrollbar(frm4, orient='vertical')
v.pack(side=tk.RIGHT, fill='y')     
lblList[16].pack(side=tk.TOP)
txtResp = tk.Text(master=frm4, yscrollcommand=v.set)
txtResp.tag_configure('e', foreground='red')
txtResp.tag_configure('s', foreground='blue')
txtResp.tag_configure('r', foreground='green')
txtResp.pack(side=tk.TOP)

# Setup Info and Checks buttons
butVer = tk.Button(text="Info", master=frm5, padx=10, pady=5)
butGfs = tk.Button(text="Driver Check", master=frm5, padx=10, pady=5)
butScm = tk.Button(text="Positioner Check", master=frm5, padx=10, pady=5)
butTxtRespClear = tk.Button(text="Clear command history", master=frm5, padx=10, pady=5)
butWinDevMan = tk.Button(text="Device Manager\n(Windows)", master=frm5)

# Layout of Info and Checks buttons frame
butTxtRespClear.grid(row=1, column=1, padx=10, pady=10)  
frm5.grid_columnconfigure(2, minsize=20)
butVer.grid(row=1, column=3, sticky=tk.W)
butGfs.grid(row=1, column=4, sticky=tk.W)
butScm.grid(row=1, column=5, sticky=tk.W)
frm5.grid_columnconfigure(6, minsize=20)
butWinDevMan.grid(row=1, column=7, sticky=tk.W)

# Layout of Disclaimer frame
lblList[17].grid(row=1, column=1)

# Functions
def butVer_handle_click(event):
   butVerThrd=thrd.Thread(target=butVer_handle_thread, daemon=True)
   butVerThrd.start()
    
def butVer_handle_thread():  
   cmdVer = '/VER'
   cmdFiv = 'FIV '
   txtResp.insert('end', ('==> Get firmware version information\n'))
   if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdVer + '\n'), 's') 
   while not oemReadQueue.empty(): 
       time.sleep(0.1)
   try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:   
           response = usbVcp.WriteRead(cmdVer, 1)
           txtResp.insert('end', ('<== CPSC1 version: ' + response + '\n'), 'r')
           for x in range(6):
               if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdFiv + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdFiv + str(x+1), 1)
               txtResp.insert('end', ('<== Slot ' + str(x+1) + ': ' + response + '\n'), 'r')            
               txtResp.see("end")
   except IOError:
       txtResp.insert('end', errConnect, 'e')      

def butGfs_handle_click(event):
    butGfsThrd=thrd.Thread(target=butGfs_handle_thread, daemon=True)
    butGfsThrd.start()
     
def butGfs_handle_thread():  
    cmdGfs = 'GFS '
    txtResp.insert('end', ('==> Get CADM2 failsafe state\n'))
    while not oemReadQueue.empty(): 
        time.sleep(0.1)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:    
           for x in range(6):
               if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdGfs + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdGfs + str(x+1), 1)
               txtResp.insert('end', ('<== Slot ' + str(x+1) + ': ' + response + '\n'), 'r')
               txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')        

def butScm_handle_click(event):
    butScmThrd=thrd.Thread(target=butScm_handle_thread, daemon=True)
    butScmThrd.start()

def butScm_handle_thread():  
    cmdScm = 'SCM '
    txtResp.insert('end', ('==> Check piezo in positioners (will take some time) ... \n'))
    parList[3][5].set(False)
    while not oemReadQueue.empty(): 
        time.sleep(0.1)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp: 
           for x in range(6):
               if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdScm + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdScm + str(x+1), 1)
               txtResp.insert('end', ('<== Positioner on slot ' + str(x+1) + ': ' + response + '[nF]\n'), 'r')
               txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')     
        
def butStp1_handle_click(event):
    butStp1Thrd=thrd.Thread(target=butStp_handle_thread, args=([1]), daemon=True)
    butStp1Thrd.start()

def butStp2_handle_click(event):
    butStp2Thrd=thrd.Thread(target=butStp_handle_thread, args=([2]), daemon=True)
    butStp2Thrd.start()
    
def butStp3_handle_click(event):
    butStp3Thrd=thrd.Thread(target=butStp_handle_thread, args=([3]), daemon=True)
    butStp3Thrd.start()
     
def butStp_handle_thread(channel): 
    cmdStp = 'STP ' + str(parList[channel-1][0].get())
    txtResp.insert('end', ('==> Stop movement\n'))
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
    while not oemReadQueue.empty(): 
        time.sleep(0.1) 
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:          
           response = usbVcp.WriteRead(cmdStp, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')       
    parList[3][6].set(False)

def butMov1_handle_click(event):
    butMov1Thrd=thrd.Thread(target=butMov_handle_thread, args=([1]), daemon=True)
    butMov1Thrd.start()

def butMov2_handle_click(event):
    butMov2Thrd=thrd.Thread(target=butMov_handle_thread, args=([2]), daemon=True)
    butMov2Thrd.start()
    
def butMov3_handle_click(event):
    butMov3Thrd=thrd.Thread(target=butMov_handle_thread, args=([3]), daemon=True)
    butMov3Thrd.start()
     
def butMov_handle_thread(channel):  
    cmdMov = 'MOV ' + str(parList[channel-1][0].get()) + ' ' + str(parList[channel-1][1].get()) + ' ' + str(parList[channel-1][2].get()) + ' ' + str(parList[channel-1][3].get()) + ' ' + str(parList[channel-1][4].get()) + ' ' + str(parList[3][1].get()) + ' ' + str(parList[channel-1][5].get()) + ' ' + str(parList[channel-1][6].get())
    txtResp.insert('end', ('==> Start movement ...\n'))
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMov + '\n'), 's')
    while not oemReadQueue.empty(): 
        time.sleep(0.1)
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMov, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
       txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')    
    parList[3][6].set(False)

def butCsz1_handle_click(event):
    butCszThrd=thrd.Thread(target=butCsz_handle_thread, args=([1]), daemon=True)
    butCszThrd.start()

#def butCsz2_handle_click(event):
#    butCszThrd=thrd.Thread(target=butCsz_handle_thread, args=([2]), daemon=True)
#    butCszThrd.start()

#def butCsz3_handle_click(event):
#    butCszThrd=thrd.Thread(target=butCsz_handle_thread, args=([3]), daemon=True)
#    butCszThrd.start()

def butCsz_handle_thread(channel):  
    cmdCsz = 'CSZ ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) 
    txtResp.insert('end', ('==> Reset counter for CH' + str(channel) + '... \n'))
    parList[3][5].set(False)
    while not oemReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdCsz + '\n'), 's')
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdCsz, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           txtResp.see("end") 
           oemList[channel][1].config(text='0', fg='black')         
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
                
def oemRead_thread():
    while(True):
        if (parList[3][5].get()) and not(parList[3][6].get()):
            oemReadQueue.put(1)
            cmdCgva = 'CGVA ' + str(parList[3][2].get())       
            try:
               with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
                   response = usbVcp.WriteRead(cmdCgva, 1)
                   responseSplit = response.split(",")                       
                   oemList[1][1].config(text=responseSplit[0], fg='black') 
                   oemList[2][1].config(text=responseSplit[1], fg='black')   
                   oemList[3][1].config(text=responseSplit[2], fg='black')                  
            except IOError:
                #txtResp.insert('end', '==> Communication briefly interrupted\n', 'e')  
                #txtResp.see("end")
                pass
            oemReadQueue.get(1)
        time.sleep(0.1)   

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
butMov1.bind("<Button-1>", butMov1_handle_click)
#butMov2.bind("<Button-1>", butMov2_handle_click)
#butMov3.bind("<Button-1>", butMov3_handle_click)
butStp1.bind("<Button-1>", butStp1_handle_click)
#butStp2.bind("<Button-1>", butStp2_handle_click)
#butStp3.bind("<Button-1>", butStp3_handle_click)
butCsz1.bind("<Button-1>", butCsz1_handle_click)
#butCsz2.bind("<Button-1>", butCsz2_handle_click)
#butCsz3.bind("<Button-1>", butCsz3_handle_click)

butTxtRespClear.bind("<Button-1>", txtResp_clear_click)
butWinDevMan.bind("<Button-1>", butWinDevMan_clear_click)

butVer.bind("<Double-Button-1>", doNothing)
butGfs.bind("<Double-Button-1>", doNothing)
butScm.bind("<Double-Button-1>", doNothing)
butMov1.bind("<Double-Button-1>", doNothing)
#butMov2.bind("<Double-Button-1>", doNothing)
#butMov3.bind("<Double-Button-1>", doNothing)
butStp1.bind("<Double-Button-1>", doNothing)
#butStp2.bind("<Double-Button-1>", doNothing)
#butStp3.bind("<Double-Button-1>", doNothing)
butCsz1.bind("<Double-Button-1>", doNothing)
#butCsz3.bind("<Double-Button-1>", doNothing)

butTxtRespClear.bind("<Double-Button-1>", doNothing)
butWinDevMan.bind("<Double-Button-1>", doNothing)

# Bind key press events to functions  
window.bind("<Insert>", butMov1_handle_click)
window.bind("<Home>", butMov2_handle_click)
window.bind("<Prior>", butMov3_handle_click)
window.bind("<Delete>", butStp1_handle_click)  
window.bind("<End>", butStp2_handle_click)  
window.bind("<Next>", butStp3_handle_click)  

# Configure and start RSM read out thread
oemReadQueue = Queue()
oemReadThrd=thrd.Thread(target=oemRead_thread)
oemReadThrd.daemon = True
oemReadThrd.start()

# Display some initial tips and hints
txtResp.insert('end', ('*** Note: make sure OEM calibration settings have been set first (CPSC1_OEM-Calibration_GUI) before starting to monitor the counter (CNT) value.\n'))
txtResp.insert('end', ('*** Remember: the OEM/COE is a relative sensor (it only counts "pulses") and that the counter value will be reset at Power OFF of the controller.\n'))

# Main loop (loop until window is closed)
window.mainloop()
