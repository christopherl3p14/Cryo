###############################################################################
# File name:      CPSC1_ServoDrive-COE_GUI_vX.y.py
# Creation:       2022-12-13
# Last Update:    2023-01-31
# Author:         JPE
# Python version: 3.9, requires pyserial and tkinter
# Description:    MVP GUI controlling CPSC1 "Servodrive" mode of operation
#                 for up to 3 positioners with feedback (3x CADM + 1x OEM)
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
window.title(f'CPSC1 ServoDrive for 3 positioners with COE feedback ({verNumber})') 
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

# Setup param list defaults
parList = [[0]*8 for i in range(4)]
parList[0][0] = tk.StringVar(window) # X-axis stage type (CADM Address 1)
parList[0][0].set('CLA2201-COE')
parList[0][1] = tk.StringVar(window) # X-axis step frequency
parList[0][1].set('600')
parList[0][2] = tk.StringVar(window) # Y-axis stage type (CADM Address 2)
parList[0][2].set('CLA2201-COE')
parList[0][3] = tk.StringVar(window) # Y-axis step frequency
parList[0][3].set('600')
parList[0][4] = tk.StringVar(window) # Z-axis stage type (CADM Address 3)
parList[0][4].set('CLA2201-COE')
parList[0][5] = tk.StringVar(window) # Z-axis step frequency
parList[0][5].set('600')

parList[1][1] = tk.IntVar() # X-axis Abs (true) or Rel (false)
parList[1][1].set(False)
parList[1][3] = tk.IntVar() # Y-axis Abs (true) or Rel (false)
parList[1][3].set(False)
parList[1][5] = tk.IntVar() # Z-axis Abs (true) or Rel (false)
parList[1][5].set(False)

parList[2][0] = tk.StringVar(window) # Drive Factor
parList[2][0].set('1')
parList[2][1] = tk.StringVar(window) # Environment temperature
parList[2][1].set('293')
parList[2][2] = tk.IntVar()          # Toggle Servodrive Status
parList[2][2].set(False)
parList[3][0] = tk.StringVar(window) # COM Port
parList[3][0].set('1')
parList[3][3] = tk.IntVar()          # Show commands toggle
parList[3][3].set(False)
parList[3][6] = tk.IntVar()          # Command lock
parList[3][6].set(False)

# Setup text labels
lblList = [0]*22
lblList[0] = tk.Label(text="COM port: ", master=frm1)
lblList[1] = tk.Label(text="Baudrate: ", master=frm1)
lblList[2] = tk.Label(text="Show Commands: ", master=frm1)
lblList[3] = tk.Label(text="Stage Type: ",master=frm2) 
lblList[4] = tk.Label(text="Freq [Hz]: ",master=frm2)
lblList[5] = tk.Label(text="CH1: ",master=frm2) 
lblList[6] = tk.Label(text="CH2: ",master=frm2) 
lblList[7] = tk.Label(text="CH3: ",master=frm2) 
lblList[12] = tk.Label(text="Setpoint [m] or [rad]: ",master=frm2) 
lblList[15] = tk.Label(text="Enabled?",master=frm2) 
lblList[16] = tk.Label(text="Finished?",master=frm2) 
lblList[17] = tk.Label(text="Invalid SP:",master=frm2) 
lblList[18] = tk.Label(text="Pos Error:",master=frm2) 

lblList[13] = tk.Label(text="Abs?: ",master=frm2)
lblList[8] = tk.Label(text="Drive Factor: ",master=frm1) 
lblList[9] = tk.Label(text="Temperature [K]: ",master=frm1) 
lblList[14] = tk.Label(text="Poll Status: ",master=frm3)
lblList[10] = tk.Label(text="Command History", master=frm4)
lblList[11] = tk.Label(text="JPE - Driven by innovation | jpe-innovations.com", master=frm6)

# Setup error messages
errConnect = 'Cannot connect to controller!\nPlease check if CPSC1 is connected to the host via USB and the correct COM port number and Baudrate has been set.\n'

# Setup user inputs
inpList = [[0]*8 for i in range(4)]
inpList[0][0] = tk.OptionMenu(frm2, parList[0][0], "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE")
inpList[0][1] = tk.Spinbox(frm2, from_= 1, to = 600, textvariable=parList[0][1], width=5)
inpList[0][2] = tk.OptionMenu(frm2, parList[0][2], "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE")
inpList[0][3] = tk.Spinbox(frm2, from_= 1, to = 600, textvariable=parList[0][3], width=5)
inpList[0][4] = tk.OptionMenu(frm2, parList[0][4], "CLA2201-COE", "CLA2601-COE", "CLA2602-COE", "CLA2603-COE", "CRM1-COE", "CLD1-COE")
inpList[0][5] = tk.Spinbox(frm2, from_= 1, to = 600, textvariable=parList[0][5], width=5)

inpList[1][0] = tk.Entry(width=8, master=frm2) # Setpoint X
inpList[1][1] = tk.Checkbutton(frm2, text = "", variable = parList[1][1], onvalue = True, offvalue = False) # Abs mode (True)
inpList[1][2] = tk.Entry(width=8, master=frm2) # Setpoint Y 
inpList[1][3] = tk.Checkbutton(frm2, text = "", variable = parList[1][3], onvalue = True, offvalue = False) # Abs mode (True)
inpList[1][4] = tk.Entry(width=8, master=frm2) # Setpoint Z 
inpList[1][5] = tk.Checkbutton(frm2, text = "", variable = parList[1][5], onvalue = True, offvalue = False) # Abs mode (True)

inpList[2][0] = tk.OptionMenu(frm1, parList[2][0], "0.6", "0.8", "1", "1.5", "2", "3") # Drive Factor
inpList[2][1] = tk.Spinbox(frm1, from_= 0, to = 300, textvariable=parList[2][1], width=5) # Temperature
inpList[2][2] = tk.Checkbutton(frm3, text = "", variable = parList[2][2], onvalue = True, offvalue = False) # Poll FBST (True)
inpList[3][0] = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=parList[3][0], width=5) # COM port
inpList[3][3] = tk.Checkbutton(frm1, text = "", variable = parList[3][3], onvalue = True, offvalue = False) # Show commands toggle
inpList[3][4] = tk.Entry(width=8, master=frm1) # BaudRate

# Set default parameter values for (Entry) widgets
inpList[1][0].insert(0, "0.0000") # Setpoint X  
inpList[1][0].config(width=10)
inpList[1][2].insert(0, "0.0000") # Setpoint Y  
inpList[1][2].config(width=10)
inpList[1][4].insert(0, "0.0000") # Setpoint Z 
inpList[1][4].config(width=10)
inpList[3][4].insert(0, "115200") # BaudRate         

# Setup Status List
stsList = [[0]*8 for i in range(4)]
stsList[0][0] = tk.Label(text="0",master=frm2) 
stsList[0][1] = tk.Label(text="0",master=frm2) 
stsList[1][0] = tk.Label(text="0",master=frm2) 
stsList[1][1] = tk.Label(text="0",master=frm2) 
stsList[1][2] = tk.Label(text="0",master=frm2) 
stsList[2][0] = tk.Label(text="0",master=frm2) 
stsList[2][1] = tk.Label(text="0",master=frm2) 
stsList[2][2] = tk.Label(text="0",master=frm2) 

# Setup Move and Stop buttons   
butFben = tk.Button(text="FBEN [Home]", master=frm3, padx=20)
butFbes = tk.Button(text="FBES [PgUp]", master=frm3, padx=20)
butFbxt = tk.Button(text="FBXT [End]", master=frm3, padx=20)
butFbcs = tk.Button(text="FBCS [PgDwn]", master=frm3, padx=20)

# Layout of user inputs frame
lblList[0].grid(row=1, column=1) # Label COM port
inpList[3][0].grid(row=1, column=2)
lblList[1].grid(row=1, column=3) # Label Baudrate
inpList[3][4].grid(row=1, column=4)
lblList[2].grid(row=1, column=5) # Label Verbose
inpList[3][3].grid(row=1, column=6)
lblList[8].grid(row=1, column=7) # Label Drive Factor
inpList[2][0].grid(row=1, column=8)
lblList[9].grid(row=1, column=9) # Label Temperature
inpList[2][1].grid(row=1, column=10)

lblList[3].grid(row=0, column=1) # Label Stage Type
lblList[4].grid(row=0, column=2) # Label Frequency
lblList[12].grid(row=0, column=3) # Label Setpoint
lblList[13].grid(row=0, column=4) # Label Abs mode
lblList[18].grid(row=0, column=5) # Label Pos Error
lblList[17].grid(row=0, column=6) # Label Invalid SP
lblList[15].grid(row=0, column=7) # Label Enabled?
lblList[16].grid(row=0, column=8) # Label Finished?

lblList[5].grid(row=1, column=0) # Label Channel 1
lblList[6].grid(row=2, column=0) # Label Channel 2
lblList[7].grid(row=3, column=0) # Label Channel 3
inpList[0][0].grid(row=1, column=1)
inpList[0][1].grid(row=1, column=2)
inpList[0][2].grid(row=2, column=1)
inpList[0][3].grid(row=2, column=2)
inpList[0][4].grid(row=3, column=1)
inpList[0][5].grid(row=3, column=2)
inpList[1][0].grid(row=1, column=3)
inpList[1][1].grid(row=1, column=4)
inpList[1][2].grid(row=2, column=3)
inpList[1][3].grid(row=2, column=4)
inpList[1][4].grid(row=3, column=3)          
inpList[1][5].grid(row=3, column=4)   
stsList[2][0].grid(row=1, column=5)
stsList[2][1].grid(row=2, column=5)
stsList[2][2].grid(row=3, column=5)
stsList[1][0].grid(row=1, column=6)
stsList[1][1].grid(row=2, column=6)
stsList[1][2].grid(row=3, column=6)
stsList[0][0].grid(row=1, column=7)
stsList[0][1].grid(row=1, column=8)

butFben.grid(row=1, column=1)
butFbes.grid(row=1, column=2)
butFbxt.grid(row=1, column=4)
lblList[14].grid(row=1, column=5) # Label Poll Status
inpList[2][2].grid(row=1, column=6)    
butFbcs.grid(row=1, column=7)
     
# Layout of Command History frame
v = tk.Scrollbar(frm4, orient='vertical')
v.pack(side=tk.RIGHT, fill='y')     
lblList[10].pack(side=tk.TOP)
txtResp = tk.Text(master=frm4, yscrollcommand=v.set)
txtResp.tag_configure('e', foreground='red')
txtResp.tag_configure('s', foreground='blue')
txtResp.tag_configure('r', foreground='green')
txtResp.pack(side=tk.TOP)

# Setup Info and Checks buttons
butVer = tk.Button(text="Info", master=frm5, padx=10, pady=5)
butTxtRespClear = tk.Button(text="Clear command history", master=frm5, padx=10, pady=5)
butWinDevMan = tk.Button(text="Device Manager\n(Windows)", master=frm5)

# Layout of Info and Checks buttons frame
butTxtRespClear.grid(row=1, column=1, padx=10, pady=10)  
frm5.grid_columnconfigure(2, minsize=20)
butVer.grid(row=1, column=3, sticky=tk.W)
frm5.grid_columnconfigure(6, minsize=20)
butWinDevMan.grid(row=1, column=7, sticky=tk.W)

# Layout of Disclaimer frame
lblList[11].grid(row=1, column=1)

# Functions
def butVer_handle_click(event):
   butVerThrd=thrd.Thread(target=butVer_handle_thread, daemon=True)
   butVerThrd.start()
    
def butVer_handle_thread():  
   cmdVer = '/VER'
   cmdFiv = 'FIV '
   txtResp.insert('end', ('==> Get firmware version information\n'))
   if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdVer + '\n'), 's') 
   while not fbstReadQueue.empty(): 
       time.sleep(0.1)
   parList[2][2].set(False)
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
           
def butFben_handle_click(event):
    butFbenThrd=thrd.Thread(target=butFben_handle_thread, daemon=True)
    butFbenThrd.start()

def butFben_handle_thread():  
    cmdFben = 'FBEN ' + str(parList[0][0].get()) + ' ' + str(parList[0][1].get()) + ' ' + str(parList[0][2].get()) + ' ' + str(parList[0][3].get()) + ' ' + str(parList[0][4].get()) + ' ' + str(parList[0][5].get()) + ' ' + str(parList[2][0].get()) + ' ' + str(parList[2][1].get())
    txtResp.insert('end', ('==> Enable Servodrive mode with set parameters ... \n'))
    while not fbstReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdFben + '\n'), 's')
    parList[2][2].set(True)
    parList[3][6].set(True)
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
            response = usbVcp.WriteRead(cmdFben, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
            txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)

def butFbxt_handle_click(event):
    butFbxtThrd=thrd.Thread(target=butFbxt_handle_thread, daemon=True)
    butFbxtThrd.start()

def butFbxt_handle_thread():  
    cmdFbxt = 'FBXT' 
    txtResp.insert('end', ('==> Disable Servodrvie mode... \n'))
    while not fbstReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdFbxt + '\n'), 's')  
    parList[3][6].set(True)
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
            response = usbVcp.WriteRead(cmdFbxt, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
            txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)

def butFbes_handle_click(event):
    butFbesThrd=thrd.Thread(target=butFbes_handle_thread, daemon=True)
    butFbesThrd.start()

def butFbes_handle_thread():  
    cmdFbes = 'FBES' 
    txtResp.insert('end', ('==> Stop current ServoDrive motion... \n'))
    while not fbstReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdFbes + '\n'), 's')  
    parList[3][6].set(True)
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
            response = usbVcp.WriteRead(cmdFbes, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
            txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)

def butFbcs_handle_click(event):
    butFbcsThrd=thrd.Thread(target=butFbcs_handle_thread, daemon=True)
    butFbcsThrd.start()

def butFbcs_handle_thread():  
    cmdFbcs = 'FBCS ' + str(inpList[1][0].get()) + ' ' + str(parList[1][1].get()) + ' ' + str(inpList[1][2].get()) + ' ' + str(parList[1][3].get()) + ' ' + str(inpList[1][4].get()) + ' ' + str(parList[1][5].get())
    txtResp.insert('end', ('==> Start moving to setpoint ... \n'))
    while not fbstReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdFbcs + '\n'), 's')
    parList[2][2].set(True)
    parList[3][6].set(True)
    try:
        with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
            response = usbVcp.WriteRead(cmdFbcs, 1)
            txtResp.insert('end', ('<== ' + response + '\n'), 'r')
            txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)
           
def fbstRead_thread():
    while(True):
        if (parList[2][2].get()) and not(parList[3][6].get()):
            fbstReadQueue.put(1)
            cmdFbst = 'FBST'
            try:
                with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
                    response = usbVcp.WriteRead(cmdFbst, 1)
                    responseSplit = response.split(",")
                    stsList[0][0].config(text=responseSplit[0], fg='green') # Enable flag
                    stsList[0][1].config(text=responseSplit[1], fg='purple') # Finished flag
                    stsList[1][0].config(text=responseSplit[2], fg='red') # Invalid SP X
                    stsList[1][1].config(text=responseSplit[3], fg='red') # Invalid SP Y
                    stsList[1][2].config(text=responseSplit[4], fg='red') # Invalid SP Z
                    stsList[2][0].config(text=responseSplit[5], fg='blue') # Pos Error X
                    stsList[2][1].config(text=responseSplit[6], fg='blue') # Pos Error Y
                    stsList[2][2].config(text=responseSplit[7], fg='blue') # Pos Error Z
                    #txtResp.insert('end', ('<== ' + response + '\n'), 'r')
                    #txtResp.see("end")    
            except IOError:
                pass
            fbstReadQueue.get(1)
        time.sleep(0.5)   

def txtResp_clear_click(event):
    txtResp.delete("1.0", "end")

def butWinDevMan_clear_click(event):
    sp.call("control /name Microsoft.DeviceManager")
    
def doNothing(event):
    time.sleep(1)

# Bind button press (left mouse click) events to functions
butFben.bind("<Button-1>", butFben_handle_click)
butFbxt.bind("<Button-1>", butFbxt_handle_click)
butFbes.bind("<Button-1>", butFbes_handle_click)
butFbcs.bind("<Button-1>", butFbcs_handle_click)
butVer.bind("<Button-1>", butVer_handle_click)
butTxtRespClear.bind("<Button-1>", txtResp_clear_click)
butWinDevMan.bind("<Button-1>", butWinDevMan_clear_click)

butFben.bind("<Double-Button-1>", doNothing)
butFbxt.bind("<Double-Button-1>", doNothing)
butFbes.bind("<Double-Button-1>", doNothing)
butFbcs.bind("<Double-Button-1>", doNothing)
butVer.bind("<Double-Button-1>", doNothing)
butTxtRespClear.bind("<Double-Button-1>", doNothing)
butWinDevMan.bind("<Double-Button-1>", doNothing)

# Bind key press events to functions  
window.bind("<Home>", butFben_handle_click)
window.bind("<End>", butFbxt_handle_click)
window.bind("<Prior>", butFbes_handle_click)
window.bind("<Next>", butFbcs_handle_click)  

fbstReadQueue = Queue()
fbstReadThrd=thrd.Thread(target=fbstRead_thread)
fbstReadThrd.daemon = True
fbstReadThrd.start()

# Display some initial tips and hints
txtResp.insert('end', ('==> Note: Before using ServoDrive make sure the connected OEM module has been calibrated for the connected -COE sensor(s). See Software User Manual for more information.\n'))

# Main loop (loop until window is closed)
window.mainloop()



