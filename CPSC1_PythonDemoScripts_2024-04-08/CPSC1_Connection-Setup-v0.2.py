###############################################################################
# File name:      CPSC1_ConnectionSetup_GUI-vX.y.py
# Creation:       2023-07-04
# Last Update:    2024-03-22
# Author:         JPE
# Python version: 3.9, requires pyserial and tkinter
# CPSC firmware:  8.0.20230516 or higher
# Description:    MVP GUI connection and communication setup of a CPSC1
# Disclaimer:     This program is provided 'As Is' without any express or 
#                 implied warranty of any kind.
###############################################################################

verNumber = 'v0.2'

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
window.title(f'CPSC1 Connection Setup ({verNumber})') 
window.option_add( "*font", "Helvetica 11" )

# Setup frames within window
frm1 = tk.Frame(padx=10, pady=10)
frm2 = tk.Frame(padx=10, pady=10)
frm3 = tk.Frame(padx=10, pady=10) 
frm4 = tk.Frame(padx=10, pady=10) 
frm5 = tk.Frame(padx=10, pady=10) 
frm1.pack() # Inputs
frm2.pack() # Main control buttons
frm3.pack() # Command history
frm4.pack() # Info and checks
frm5.pack() # Disclaimer

# Setup text labels
lblList = [0]*50
lblList[0] = tk.Label(text="COM port (USB): ", master=frm1, width=15, anchor='e')                                   # COM port
lblList[1] = tk.Label(text="Baudrate (USB): ", master=frm1, width=15, anchor='e')                                   # Baudrate
lblList[2] = tk.Label(text="Baudrate (RS422): ", master=frm1, width=15, anchor='e')                                 # Baudrate
lblList[3] = tk.Label(text="Show Commands: ", master=frm1, width=15, anchor='e')                                    # Show Commands
lblList[4] = tk.Label(text="DHCP: ", master=frm1, width=15, anchor='e')                                             # DHCP
lblList[5] = tk.Label(text="IP Address: ", master=frm1, width=15, anchor='e')                                       # IP Address
lblList[6] = tk.Label(text="Subnet Mask: ",master=frm1, width=15, anchor='e')                                       # Subnet Mask
lblList[7] = tk.Label(text="Gateway: ",master=frm1, width=15, anchor='e')                                           # Gateway 
lblList[8] = tk.Label(text="Baudrate (USB): ", master=frm2, anchor='e')                                             # Baudrate Parameters USB / Virtual COM
lblList[9] = tk.Label(text="Baudrate (RS422): ", master=frm2, anchor='e')                                           # Baudrate Parameters RS422
lblList[10] = tk.Label(text="LAN Parameters: ", master=frm2, anchor='e')                                            # LAN Parameters
lblList[11] = tk.Label(text="Command History", master=frm3, anchor='center')                                        # Command History
lblList[12] = tk.Label(text="JPE - Driven by innovation | jpe-innovations.com", master=frm4, anchor='center')       # Copyright info
lblList[13] = tk.Label(text="MAC Address: ",master=frm1, width=15, anchor='e')                                      # MAC Address


# Setup error messages
errConnect = 'Cannot connect to controller!\nPlease check if CPSC1 is connected to the host via USB and the correct COM port number and Baudrate has been set.\n'

# Set default parameter values (Spinbox and OptionMenu)
optList = [0]*50
optList[0] = tk.StringVar(window)                                                                       # Set default value for COM port
optList[0].set('1')                                                                                     # Set this to 1
optList[2] = tk.IntVar()                                                                                # Set default value for Show Commands checkbox
optList[2].set(True)                                                                                    # Set this to TRUE
optList[3] = tk.IntVar()                                                                                # Set default value for DHCP checkbox
optList[3].set(True)                                                                                    # Set this to TRUE

# Setup user inputs (frm1)
inpList = [0]*50
inpList[0] = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=optList[0], width=6)                     # COM Port input spinbox
inpList[1] = tk.Entry(width=8, master=frm1)                                                             # Baurate (USB) input field
inpList[1].insert(0, "115200")                                                                          # Insert default value for Baudrate (USB) input field
inpList[2] = tk.Entry(width=8, master=frm1)                                                             # Baurate (RS422) input field
inpList[2].insert(0, "115200")                                                                          # Insert default value for Baudrate (RS422) input field
inpList[3] = tk.Checkbutton(frm1, text = "", variable = optList[2], onvalue = True, offvalue = False)   # Show Commands checkbox
inpList[4] = tk.Checkbutton(frm1, text = "", variable = optList[3], onvalue = True, offvalue = False)   # DHCP checkbox
inpList[5] = tk.Entry(width=16, master=frm1)                                                            # IP Address input field
inpList[5].insert(0, "0.0.0.0")                                                                         # Insert default value for IP Address input field
inpList[6] = tk.Entry(width=16, master=frm1)                                                            # Subnet Mask input field
inpList[6].insert(0, "255.255.255.0")                                                                   # Insert default value for Subnet Mask input field
inpList[7] = tk.Entry(width=16, master=frm1)                                                            # Gateway input field
inpList[7].insert(0, "0.0.0.0")                                                                         # Insert default value for Gateway input field
inpList[13] = tk.Entry(width=16, master=frm1)                                                           # Gateway input field
inpList[13].insert(0, "[To be Retreived]")                                                              # MAC Address

# Buttons (frm2)
butList = [0]*50
butList[0] = tk.Button(text="Info", master=frm2, padx=10, pady=5)                                       # Version Info button
butList[1] = tk.Button(text="Get", master=frm2, padx=10, pady=5)                                        # Get Baudrate (USB) Settings button
butList[2] = tk.Button(text="Set", master=frm2, padx=10, pady=5)                                        # Set Baudrate (USB) Settings button
butList[3] = tk.Button(text="Get", master=frm2, padx=10, pady=5)                                        # Get Baudrate (RS422) Settings button
butList[4] = tk.Button(text="Set", master=frm2, padx=10, pady=5)                                        # Set Baudrate (RS422) Settings button
butList[5] = tk.Button(text="Get", master=frm2, padx=10, pady=5)                                        # Get LAN Settings button
butList[6] = tk.Button(text="Set", master=frm2, padx=10, pady=5)                                        # Set LAN Settings button
butList[8] = tk.Button(text="Clear command history", master=frm2, padx=10, pady=5)                      # Clear Command History button
butList[9] = tk.Button(text="Device Manager (Windows)", master=frm2, padx=10, pady=5)                   # Open Device Manager button
butList[10] = tk.Button(text=">> How To Use <<", master=frm2, padx=10, pady=5)                          # Display Help

# Layout of FRAME 1 (frm1)
lblList[0].grid(row=1, column=1, sticky=tk.E)                                                           # COM Port (USB) Label
inpList[0].grid(row=1, column=2, sticky=tk.W)                                                           # COM Port (USB) Input
lblList[1].grid(row=2, column=1, sticky=tk.E)                                                           # Baudrate (USB) Label
inpList[1].grid(row=2, column=2, sticky=tk.W)                                                           # Baurdate (USB) Input
lblList[2].grid(row=3, column=1, sticky=tk.E)                                                           # Baudrate (RS422) Label
inpList[2].grid(row=3, column=2, sticky=tk.W)                                                           # Baurdate (RS422) Input
lblList[3].grid(row=4, column=1, sticky=tk.E)                                                           # Show Commands Label
inpList[3].grid(row=4, column=2, sticky=tk.W)                                                           # Show Commands Input

lblList[4].grid(row=1, column=3, sticky=tk.E)                                                           # DHCP Label
inpList[4].grid(row=1, column=4, sticky=tk.W)                                                           # DHCP Input
lblList[5].grid(row=2, column=3, sticky=tk.E)                                                           # IP Address Label
inpList[5].grid(row=2, column=4, sticky=tk.W)                                                           # IP Address Input
lblList[6].grid(row=3, column=3, sticky=tk.E)                                                           # Subnet Mask Label
inpList[6].grid(row=3, column=4, sticky=tk.W)                                                           # Subnet Mask Input
lblList[7].grid(row=4, column=3, sticky=tk.E)                                                           # Gateway Label
inpList[7].grid(row=4, column=4, sticky=tk.W)                                                           # Gateway Input
lblList[13].grid(row=5, column=3, sticky=tk.E)                                                          # MAC Address Label
inpList[13].grid(row=5, column=4, sticky=tk.W)                                                          # MAC Address Input

# Layout of FRAME 2 (frm2)
lblList[8].grid(row=1, column=1, sticky=tk.E)                                                           # LABEL: Baurate (USB) Parameters
butList[1].grid(row=1, column=2, sticky=tk.W)                                                           # BUTTON: Get
butList[2].grid(row=1, column=3, sticky=tk.W)                                                           # BUTTON: Set
lblList[9].grid(row=2, column=1, sticky=tk.E)                                                           # LABEL: Baurate (RS422) Parameters
butList[3].grid(row=2, column=2, sticky=tk.W)                                                           # BUTTON: Get
butList[4].grid(row=2, column=3, sticky=tk.W)                                                           # BUTTON: Set
lblList[10].grid(row=3, column=1, sticky=tk.E)                                                          # LABEL: LAN Parameters
butList[5].grid(row=3, column=2, sticky=tk.W)                                                           # BUTTON: Get
butList[6].grid(row=3, column=3, sticky=tk.W)                                                           # BUTTON: Set 
frm2.grid_columnconfigure(4, minsize=40)
butList[0].grid(row=1, column=5, sticky=tk.E)                                                           # BUTTON: Info
butList[8].grid(row=1, column=6, sticky=tk.E)                                                           # BUTTON: Clear Command History
butList[9].grid(row=2, column=5, columnspan=2, sticky=tk.W)                                             # BUTTON: Device Manager
butList[10].grid(row=3, column=5, columnspan=2, sticky=tk.W)                                            # BUTTON: Device Manager

                  
# Layout of FRAME 3 (frm3)
v = tk.Scrollbar(frm3, orient='vertical')                                                               # Setup scrollbar for frm3
v.pack(side=tk.RIGHT, fill='y')                                                                         # Setup scrollbar for frm3
lblList[11].pack(side=tk.TOP)                                                                           # LABEL: Command History
txtRespFont = ("Courier New", 11)                                                                       # Set specific font for TEXTBOX
txtResp = tk.Text(height=15, width=60, master=frm3, yscrollcommand=v.set)                             # Setup TEXTBOX
txtResp.tag_configure('e', foreground='red')                                                            # Setup text in TEXTBOX
txtResp.tag_configure('s', foreground='blue')                                                           # Setup text in TEXTBOX
txtResp.tag_configure('r', foreground='green')                                                          # Setup text in TEXTBOX
txtResp.configure(font = txtRespFont)                                                                   # Set specific font for TEXTBOX
txtResp.pack(side=tk.TOP)                                                                               # Layout of TEXTBOX

# Layout of FRAME 4 (frm4)
lblList[12].grid(row=1, column=1)                                                                       # LABEL: Copyright info


# ---------
# Functions
# ---------
def butVer_handle_click(event):
   butVerThrd=thrd.Thread(target=butVer_handle_thread)
   butVerThrd.start()
    
def butVer_handle_thread():  
   cmdVer = '/VER'
   cmdFiv = 'FIV '
   txtResp.insert('end', ('==> Get firmware version information\n'))
   if (optList[2].get()): txtResp.insert('end', ('==> ' + cmdVer + '\n'), 's') 
   try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optList[0].get())), str(inpList[1].get())) as usbVcp:   
           response = usbVcp.WriteRead(cmdVer, 1)
           txtResp.insert('end', ('<== CPSC1 version: ' + response + '\n'), 'r')
           for x in range(6):
               if (optList[2].get()): txtResp.insert('end', ('==> ' + cmdFiv + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdFiv + str(x+1), 1)
               txtResp.insert('end', ('<== Slot ' + str(x+1) + ': ' + response + '\n'), 'r')            
               txtResp.see("end")
   except IOError:
       txtResp.insert('end', errConnect, 'e')      

def butBaudrateGetUsb(event):
    butBaudrateThrd=thrd.Thread(target=butBaudrateGet, args=(['USB']), daemon=True)
    butBaudrateThrd.start()

def butBaudrateGetRs422(event):
    butBaudrateThrd=thrd.Thread(target=butBaudrateGet, args=(['RS422']), daemon=True)
    butBaudrateThrd.start()
     
def butBaudrateGet(interface):
    cmdGbr = '/GBR ' + str(interface)
    txtResp.insert('end', ('==> Get current Baudrate parameters\n'))
    if (optList[2].get()): txtResp.insert('end', ('==> ' + cmdGbr + '\n'), 's')
 
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optList[0].get())), str(inpList[1].get())) as usbVcp:           
           response = usbVcp.WriteRead(cmdGbr, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           txtResp.see("end")
           if (interface == 'USB'):
               inpList[1].delete(0 ,'end')
               inpList[1].insert(0, response)  
           else:
               inpList[2].delete(0 ,'end')
               inpList[2].insert(0, response) 
    except IOError:
        txtResp.insert('end', errConnect, 'e')       

def butBaudrateSetUsb(event):
    butBaudrateThrd=thrd.Thread(target=butBaudrateSet, args=(['USB']), daemon=True)
    butBaudrateThrd.start()

def butBaudrateSetRs422(event):
    butBaudrateThrd=thrd.Thread(target=butBaudrateSet, args=(['RS422']), daemon=True)
    butBaudrateThrd.start()
     
def butBaudrateSet(interface):
    if (interface == 'USB'):
        cmdSbr = '/SBR ' + str(interface) + ' ' + str(inpList[1].get())
    else:
        cmdSbr = '/SBR ' + str(interface) + ' ' + str(inpList[2].get())  
    txtResp.insert('end', ('==> Set Baudrate parameters\n'))
    if (optList[2].get()): txtResp.insert('end', ('==> ' + cmdSbr + '\n'), 's')
 
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optList[0].get())), str(inpList[1].get())) as usbVcp:           
           response = usbVcp.WriteRead(cmdSbr, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')      

def butIpGet(event):
    funcIpGetThrd=thrd.Thread(target=funcIpGet, daemon=True)
    funcIpGetThrd.start()

def funcIpGet():  
    cmdIpr = '/IPR '
    txtResp.insert('end', ('==> Get current IP parameters\n'))
    if (optList[2].get()): txtResp.insert('end', ('==> ' + cmdIpr + '\n'), 's')
 
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optList[0].get())), str(inpList[1].get())) as usbVcp:    
           response = usbVcp.WriteRead(cmdIpr, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           txtResp.see("end")
           responseSplit = response.split(",")
           if (responseSplit[0] == 'dhcp'):
               optList[3].set(True)  
           else:
               optList[3].set(False)  
           inpList[5].delete(0 ,'end')
           inpList[5].insert(0, responseSplit[1])  
           inpList[6].delete(0 ,'end')
           inpList[6].insert(0, responseSplit[2])
           inpList[7].delete(0 ,'end')
           inpList[7].insert(0, responseSplit[3])
           inpList[13].delete(0 ,'end')
           inpList[13].insert(0, responseSplit[4])
    except IOError:
        txtResp.insert('end', errConnect, 'e')      

def butIpSet(event):
    funcIpSetThrd=thrd.Thread(target=funcIpSet, daemon=True)
    funcIpSetThrd.start()

def funcIpSet():  
    if (optList[3].get()): cmdIps = '/IPS DHCP 0.0.0.0 0.0.0.0 0.0.0.0'
    else: cmdIps = '/IPS STATIC ' + str(inpList[5].get()) + ' ' + str(inpList[6].get()) + ' ' + str(inpList[7].get())
    txtResp.insert('end', ('==> Set current IP parameters\n'))
    if (optList[2].get()): txtResp.insert('end', ('==> ' + cmdIps + '\n'), 's')
 
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(optList[0].get())), str(inpList[1].get())) as usbVcp:    
           response = usbVcp.WriteRead(cmdIps, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           txtResp.see("end")
    except IOError:
        txtResp.insert('end', errConnect, 'e')   

def txtResp_clear_click(event):
    txtResp.delete("1.0", "end")

def butWinDevMan_clear_click(event):
    sp.call("control /name Microsoft.DeviceManager")

def doNothing(event):
    time.sleep(1)

def butHelp(event):
    # Display some initial tips and hints
    txtResp.insert('end', ('GENERAL SETTINGS\n'))
    txtResp.insert('end', ('This GUI requires the CPSC1 controller to be connected via USB (Virtual COM), so set COM Port (USB) and Baudrate (USB) first!\n'))
    txtResp.insert('end', ('\nFACTORY DEFAULT VALUES\n'))
    txtResp.insert('end', ('Baudrate (USB) : 115200\n'))
    txtResp.insert('end', ('Baudrate (RS422) : 115200\n'))
    txtResp.insert('end', ('DHCP : Enabled\n'))
    txtResp.insert('end', ('\nIMPORTANT\n'))
    txtResp.insert('end', ('Be careful when updating the Baudrate (USB) value. Make sure to remember the correct setting!\n'))
    txtResp.insert('end', ('Uncheck the "DHCP" box to set a static IP Address, Subnet Mask and Gateway.\n'))
    txtResp.insert('end', ('When using the "Get" buttons, all input boxes will be updated with retreived values.\n'))
    

# Bind button press (left mouse click) events to functions
butList[0].bind("<Button-1>", butVer_handle_click)
butList[8].bind("<Button-1>", txtResp_clear_click)
butList[9].bind("<Button-1>", butWinDevMan_clear_click)
butList[10].bind("<Button-1>", butHelp)

butList[1].bind("<Button-1>", butBaudrateGetUsb)
butList[2].bind("<Button-1>", butBaudrateSetUsb)
butList[3].bind("<Button-1>", butBaudrateGetRs422)
butList[4].bind("<Button-1>", butBaudrateSetRs422)
butList[5].bind("<Button-1>", butIpGet)
butList[6].bind("<Button-1>", butIpSet)

butList[1].bind("<Double-Button-1>", doNothing)
butList[2].bind("<Double-Button-1>", doNothing)
butList[3].bind("<Double-Button-1>", doNothing)
butList[4].bind("<Double-Button-1>", doNothing)
butList[5].bind("<Double-Button-1>", doNothing)
butList[6].bind("<Double-Button-1>", doNothing)


# ---------------------------------------
# Main loop (loop until window is closed)
# ---------------------------------------
window.mainloop()
