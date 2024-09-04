###############################################################################
# File name:      CPSC1_BaseDrive-RLS_GUI_vX.y.py
# Creation:       2022-12-07
# Last Update:    2024-03-22
# Author:         JPE
# Python version: 3.9, requires pyserial and tkinter
# Description:    MVP GUI controlling CPSC1 "Basedrive" mode of operation
#                 for up to 3 positioners with RLS (3x CADM + 1x RSM). 
#                 Including a basic "sequence" function to move back and forth
#                 between two setpoints for a set number of times.
# Disclaimer:     This program is provided 'As Is' without any express or 
#                 implied warranty of any kind.
###############################################################################

verNumber = 'v0.6'

# Defaults for drop down lists
lstStages = ['CBS5-RLS','CBS5-HF-RLS','CBS10-RLS', 'CS021-RLS.X', 'CS021-RLS.Y', 'CS021-RLS.Z', 'CLD2-RLS', 'CSR1-RRS', 'CSR2-RRS']
lstDf = ['0.5','0.6','0.7','0.8','0.9','1.0','1.1','1.2','1.3','1.4','1.5','2.0','2.5','3.0']
lstCh = ['1','2','3','4','5','6']
lstDir = ['0','1']

# 3rd party imports
import time
import tkinter as tk
import csv
from matplotlib import pyplot as plt
import subprocess as sp
import threading as thrd
from queue import Queue
import sys

# JPE imports
from CpscInterfaces import CpscSerialInterface as CpscSerial

# Create GUI window
window = tk.Tk()
 
# Setup icon global window parameters
p1 = tk.PhotoImage(file = 'jpe.png')
window.iconphoto(False, p1)
window.title(f'CPSC1 BaseDrive for 3 positioners with RLS ({verNumber})') 
window.option_add( "*font", "Helvetica 11" )
#window.geometry("600x900")

# Setup frames within window
frm1 = tk.Frame(padx=10, pady=5, borderwidth=1, relief='groove')
frm2 = tk.Frame(padx=10, pady=5, borderwidth=1, relief='groove')
frm3 = tk.Frame(padx=10, pady=5) 
frm4 = tk.Frame(padx=10, pady=5) 
frm5 = tk.Frame(padx=10, pady=5) 
frm6 = tk.Frame(padx=10, pady=5) 
frm1.pack(padx=10, pady=5) # General Inputs
frm2.pack(padx=10, pady=5) # Feedback Inputs
frm3.pack(padx=10, pady=5) # Basedrive buttons
frm4.pack(padx=10, pady=5) # Command history
frm5.pack(padx=10, pady=5) # Info and checks
frm6.pack(padx=10, pady=5) # Disclaimer

# # Setup param list defaults
parList = [[0]*11 for i in range(4)]
parList[0][0] = tk.StringVar(window) # Channel 1, CADM address
parList[0][0].set('1')
parList[1][0] = tk.StringVar(window) # Channel 2, CADM address
parList[1][0].set('2')
parList[2][0] = tk.StringVar(window) # Channel 3, CADM address
parList[2][0].set('3')
parList[0][1] = tk.StringVar(window) # Channel 1, Direction of movement
parList[0][1].set('1')
parList[1][1] = tk.StringVar(window) # Channel 2, Direction of movement
parList[1][1].set('1')
parList[2][1] = tk.StringVar(window) # Channel 3, Direction of movement
parList[2][1].set('1')
parList[0][2] = tk.StringVar(window) # Channel 1, Frequency of movement
parList[0][2].set('600')
parList[1][2] = tk.StringVar(window) # Channel 2, Frequency of movement
parList[1][2].set('600')
parList[2][2] = tk.StringVar(window) # Channel 3, Frequency of movement
parList[2][2].set('600')
parList[0][3] = tk.StringVar(window) # Channel 1, Relative Step Size of movement
parList[0][3].set('100')
parList[1][3] = tk.StringVar(window) # Channel 2, Relative Step Size of movement
parList[1][3].set('100')
parList[2][3] = tk.StringVar(window) # Channel 3, Relative Step Size of movement
parList[2][3].set('100')
parList[0][4] = tk.StringVar(window) # Channel 1, Number of steps of movement
parList[0][4].set('0')
parList[1][4] = tk.StringVar(window) # Channel 2, Number of steps of movement
parList[1][4].set('0')
parList[2][4] = tk.StringVar(window) # Channel 3, Number of steps of movement
parList[2][4].set('0')
parList[0][5] = tk.StringVar(window) # Channel 1, Connected stage type (movement)
parList[0][5].set('CS021-RLS.X')
parList[1][5] = tk.StringVar(window) # Channel 2, Connected stage type (movement)
parList[1][5].set('CS021-RLS.Y')
parList[2][5] = tk.StringVar(window) # Channel 3, Connected stage type (movement)
parList[2][5].set('CS021-RLS.Z')
parList[0][6] = tk.StringVar(window) # Channel 1, Drive Factor for movement
parList[0][6].set('1.0')
parList[1][6] = tk.StringVar(window) # Channel 2, Drive Factor for movement
parList[1][6].set('1.0')
parList[2][6] = tk.StringVar(window) # Channel 3, Drive Factor for movement
parList[2][6].set('1.0')
parList[3][0] = tk.StringVar(window) # COM Port
parList[3][0].set('7')
parList[3][1] = tk.StringVar(window) # Environment temperature
parList[3][1].set('293')
parList[3][2] = tk.StringVar(window) # RSM Address
parList[3][2].set('2')
parList[3][3] = tk.IntVar()          # Show commands toggle
parList[3][3].set(True)
parList[3][5] = tk.IntVar()          # RLS toggle
parList[3][5].set(False)
parList[3][6] = tk.IntVar()          # Command lock
parList[3][6].set(False)
parList[0][7] = tk.StringVar(window) # Channel 1, Connected stage type (RLS)
parList[0][7].set('CS021-RLS.X')
parList[1][7] = tk.StringVar(window) # Channel 2, Connected stage type (RLS)
parList[1][7].set('CS021-RLS.Y')
parList[2][7] = tk.StringVar(window) # Channel 3, Connected stage type (RLS)
parList[2][7].set('CS021-RLS.Z')

parList[0][8] = tk.StringVar(window) # Channel 1, Move Sequence # or Runs
parList[0][8].set('4')
parList[1][8] = tk.StringVar(window) # Channel 2, Move Sequence # or Runs
parList[1][8].set('4')
parList[2][8] = tk.StringVar(window) # Channel 3, Move Sequence # or Runs
parList[2][8].set('4')
parList[0][9] = tk.StringVar(window) # Channel 1, Run Sequence #
parList[0][9].set('1')
parList[1][9] = tk.StringVar(window) # Channel 2, Run Sequence #
parList[1][9].set('1')
parList[2][9] = tk.StringVar(window) # Channel 3, Run Sequence #
parList[2][9].set('1')
parList[0][10] = tk.StringVar(window) # Channel 1, Run Sequence Timeout
parList[0][10].set('10')
parList[1][10] = tk.StringVar(window) # Channel 2, Run Sequence Timeout
parList[1][10].set('10')
parList[2][10] = tk.StringVar(window) # Channel 3, Run Sequence Timeout
parList[2][10].set('10')

# Setup text labels
lblList = [0]*22
lblList[0] = tk.Label(text="COM port: ", master=frm1)
lblList[1] = tk.Label(text="Baudrate: ", master=frm1)
lblList[2] = tk.Label(text="Show Commands: ", master=frm1)
lblList[3] = tk.Label(text="RSM Address: ", master=frm1)
lblList[4] = tk.Label(text="CADM Addr: ", master=frm1)
lblList[5] = tk.Label(text="Direction: ",master=frm1)
lblList[6] = tk.Label(text="Freq [Hz]: ",master=frm1)
lblList[7] = tk.Label(text="Rss [%]: ",master=frm1)  
lblList[8] = tk.Label(text="[#] Steps: ",master=frm1) 
lblList[9] = tk.Label(text="Temperature [K]: ",master=frm1) 
lblList[10] = tk.Label(text="Stage Type: ",master=frm1) 
lblList[11] = tk.Label(text="Drive Factor: ",master=frm1) 
lblList[12] = tk.Label(text="CH1: ",master=frm1) 
lblList[13] = tk.Label(text="CH2: ",master=frm1) 
lblList[14] = tk.Label(text="CH3: ",master=frm1) 
lblList[16] = tk.Label(text="Command History", master=frm4)
lblList[17] = tk.Label(text="JPE - Driven by innovation | jpe-innovations.com", master=frm6)
lblList[18] = tk.Label(text="Positioner ID [#]: ",master=frm1)
lblList[19] = tk.Label(text="GENERAL & DRIVE SETTINGS", master=frm1)
lblList[20] = tk.Label(text="FEEDBACK SETTINGS", master=frm2)
lblList[21] = tk.Label(text="SEQUENCE RUN", master=frm2)

# Setup RLS/RSM data list
rsmList = [[0]*6 for i in range(4)]
rsmList[0][0] = tk.Label(text="RSM Polling: ",master=frm2) 
rsmList[0][1] = tk.Label(text="Curr. Pos:",master=frm2)
rsmList[0][2] = tk.Label(text="MIR:",master=frm2)
rsmList[0][3] = tk.Label(text="MAR:",master=frm2)
rsmList[0][4] = tk.Label(text="Center:",master=frm2)
rsmList[1][0] = tk.Label(text="CH1 (A): ",master=frm2) 
rsmList[2][0] = tk.Label(text="CH2 (B): ",master=frm2) 
rsmList[3][0] = tk.Label(text="CH3 (C): ",master=frm2) 
rsmList[1][1] = tk.Label(text="0.000000000",master=frm2) 
rsmList[2][1] = tk.Label(text="0.000000000",master=frm2) 
rsmList[3][1] = tk.Label(text="0.000000000",master=frm2) 
rsmList[1][2] = tk.Label(text="0.000000000",master=frm2) 
rsmList[2][2] = tk.Label(text="0.000000000",master=frm2) 
rsmList[3][2] = tk.Label(text="0.000000000",master=frm2) 
rsmList[1][3] = tk.Label(text="0.000000000",master=frm2) 
rsmList[2][3] = tk.Label(text="0.000000000",master=frm2) 
rsmList[3][3] = tk.Label(text="0.000000000",master=frm2)
rsmList[1][4] = tk.Label(text="0.00",master=frm2) 
rsmList[2][4] = tk.Label(text="0.00",master=frm2) 
rsmList[3][4] = tk.Label(text="0.00",master=frm2)

# Setup error messages
errConnect = 'Cannot connect to controller!\nPlease check if CPSC1 is connected to the host via USB and the correct COM port number and Baudrate has been set.\n'

# Setup user inputs
inpList = [[0]*9 for i in range(4)]
inpList[0][0] = tk.OptionMenu(frm1, parList[0][0], *lstCh)
inpList[1][0] = tk.OptionMenu(frm1, parList[1][0], *lstCh)
inpList[2][0] = tk.OptionMenu(frm1, parList[2][0], *lstCh)
inpList[0][1] = tk.OptionMenu(frm1, parList[0][1], *lstDir)
inpList[1][1] = tk.OptionMenu(frm1, parList[1][1], *lstDir)
inpList[2][1] = tk.OptionMenu(frm1, parList[2][1], *lstDir)
inpList[0][2] = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=parList[0][2], width=5)
inpList[1][2] = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=parList[1][2], width=5)
inpList[2][2] = tk.Spinbox(frm1, from_= 1, to = 600, textvariable=parList[2][2], width=5)
inpList[0][3] = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=parList[0][3], width=5)
inpList[1][3] = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=parList[1][3], width=5)
inpList[2][3] = tk.Spinbox(frm1, from_= 1, to = 100, textvariable=parList[2][3], width=5)
inpList[0][4] = tk.Entry(width=6, master=frm1)
inpList[1][4] = tk.Entry(width=6, master=frm1)
inpList[2][4] = tk.Entry(width=6, master=frm1)
inpList[0][5] = tk.OptionMenu(frm1, parList[0][5], *lstStages)
inpList[1][5] = tk.OptionMenu(frm1, parList[1][5], *lstStages)
inpList[2][5] = tk.OptionMenu(frm1, parList[2][5], *lstStages)
inpList[0][6] = tk.OptionMenu(frm1, parList[0][6], *lstDf)
inpList[1][6] = tk.OptionMenu(frm1, parList[1][6], *lstDf)
inpList[2][6] = tk.OptionMenu(frm1, parList[2][6], *lstDf)
inpList[0][7] = tk.OptionMenu(frm2, parList[0][7], *lstStages)
inpList[1][7] = tk.OptionMenu(frm2, parList[1][7], *lstStages)
inpList[2][7] = tk.OptionMenu(frm2, parList[2][7], *lstStages)
inpList[3][0] = tk.Spinbox(frm1, from_= 0, to = 100, textvariable=parList[3][0], width=5)
inpList[3][1] = tk.Spinbox(frm1, from_= 0, to = 300, textvariable=parList[3][1], width=5)
inpList[3][2] = tk.OptionMenu(frm1, parList[3][2], *lstCh)
inpList[3][3] = tk.Checkbutton(frm1, text = "", variable = parList[3][3], onvalue = True, offvalue = False)
inpList[3][4] = tk.Entry(width=8, master=frm1)
inpList[3][5] = tk.Checkbutton(frm2, text = "", variable = parList[3][5], onvalue = True, offvalue = False)
inpList[0][8] = tk.Entry(width=6, master=frm1)
inpList[1][8] = tk.Entry(width=6, master=frm1)
inpList[2][8] = tk.Entry(width=6, master=frm1)

# Set default parameter values for (Entry) widgets
inpList[0][4].insert(0, "0") # CH1 No of Steps
inpList[1][4].insert(0, "0") # CH2 No of Steps
inpList[2][4].insert(0, "0") # CH3 No of Steps
inpList[3][4].insert(0, "115200") # BaudRate      
inpList[0][8].insert(0, "XXYZZAAA") # Positioner ID connected to CH1
inpList[1][8].insert(0, "XXYZZAAA") # Positioner ID connected to CH2
inpList[2][8].insert(0, "XXYZZAAA") # Positioner ID connected to CH3   

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
inpList[0][7].config(width=12)
inpList[1][7].config(width=12)
inpList[2][7].config(width=12)
inpList[0][8].config(width=10)
inpList[1][8].config(width=10)
inpList[2][8].config(width=10)

# Move Sequence List
seqList = [[0]*8 for i in range(7)]
seqList[0][1] = tk.Label(text="Pos A: ",master=frm2) 
seqList[0][2] = tk.Label(text="Pos B: ",master=frm2) 
seqList[0][3] = tk.Label(text="# Runs: ",master=frm2) 
seqList[0][4] = tk.Label(text="Seq #: ",master=frm2) 
seqList[0][5] = tk.Label(text="Timeout [s]: ",master=frm2) 
seqList[0][6] = tk.Label(text="Timer [s]: ",master=frm2) 
seqList[1][1] = tk.Entry(width=8, master=frm2) # Setpoint POS A CH1
seqList[2][1] = tk.Entry(width=8, master=frm2) # Setpoint POS A CH2
seqList[3][1] = tk.Entry(width=8, master=frm2) # Setpoint POS A CH3
seqList[1][2] = tk.Entry(width=8, master=frm2) # Setpoint POS B CH1
seqList[2][2] = tk.Entry(width=8, master=frm2) # Setpoint POS B CH2
seqList[3][2] = tk.Entry(width=8, master=frm2) # Setpoint POS B CH3
seqList[1][3] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[0][8], width=5) # Number of Runs CH1
seqList[2][3] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[1][8], width=5) # Number of Runs CH2
seqList[3][3] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[2][8], width=5) # Number of Runs CH3
seqList[1][4] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[0][9], width=5) # Sequence Number CH1
seqList[2][4] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[1][9], width=5) # Sequence Number CH2
seqList[3][4] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[2][9], width=5) # Sequence Number CH3
seqList[1][5] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[0][10], width=5) # TimeOut CH1
seqList[2][5] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[1][10], width=5) # TimeOut CH2
seqList[3][5] = tk.Spinbox(frm2, from_= 1, to = 100, textvariable=parList[2][10], width=5) # TimeOut CH3
seqList[1][6] = tk.Label(text="0.0",master=frm2) # Time Passed value CH1
seqList[2][6] = tk.Label(text="0.0",master=frm2) # Time Passed value CH2
seqList[3][6] = tk.Label(text="0.0",master=frm2) # Time Passed value CH3

# Set default parameter values for (Entry) widgets
seqList[1][1].insert(0, "-0.001") # Setpoint POS A CH1
seqList[2][1].insert(0, "-0.001") # Setpoint POS A CH2
seqList[3][1].insert(0, "-0.001") # Setpoint POS A CH3
seqList[1][2].insert(0, "0.001") # Setpoint POS B CH1
seqList[2][2].insert(0, "0.001") # Setpoint POS B CH2
seqList[3][2].insert(0, "0.001") # Setpoint POS B CH3

# Setup Move and Stop buttons   
butMov1 = tk.Button(text="Move [Ins]", master=frm1, width=8, padx=10)
butMov2 = tk.Button(text="Move [Hom]", master=frm1, width=8, padx=10)
butMov3 = tk.Button(text="Move [PgU]", master=frm1, width=8, padx=10)
butStp1 = tk.Button(text="Stop [Del]", master=frm1, width=8, padx=10)
butStp2 = tk.Button(text="Stop [End]", master=frm1, width=8, padx=10)
butStp3 = tk.Button(text="Stop [PgD]", master=frm1, width=8, padx=10)

# Setup RSM config buttons
butMir1 = tk.Button(text="MIR", master=frm2, width=2, padx=10)
butMir2 = tk.Button(text="MIR", master=frm2, width=2, padx=10)
butMir3 = tk.Button(text="MIR", master=frm2, width=2, padx=10)
butMar1 = tk.Button(text="MAR", master=frm2, width=2, padx=10)
butMar2 = tk.Button(text="MAR", master=frm2, width=2, padx=10)
butMar3 = tk.Button(text="MAR", master=frm2, width=2, padx=10)
butCenter1 = tk.Button(text="Center", master=frm2, width=4, padx=10)
butCenter2 = tk.Button(text="Center", master=frm2, width=4, padx=10)
butCenter3 = tk.Button(text="Center", master=frm2, width=4, padx=10)
butMis1 = tk.Button(text="MIS", master=frm2, width=2, padx=10)
butMis2 = tk.Button(text="MIS", master=frm2, width=2, padx=10)
butMis3 = tk.Button(text="MIS", master=frm2, width=2, padx=10)
butMas1 = tk.Button(text="MAS", master=frm2, width=2, padx=10)
butMas2 = tk.Button(text="MAS", master=frm2, width=2, padx=10)
butMas3 = tk.Button(text="MAS", master=frm2, width=2, padx=10)
butMmr1 = tk.Button(text="Reset", master=frm2, width=4, padx=10)
butMmr2 = tk.Button(text="Reset", master=frm2, width=4, padx=10)
butMmr3 = tk.Button(text="Reset", master=frm2, width=4, padx=10)
butRss = tk.Button(text="Store Settings", master=frm2, width=8, padx=20)

# Setup Sequence buttons   
butRunSeq1 = tk.Button(text="Run", master=frm2, width=4, padx=10)
butRunSeq2 = tk.Button(text="Run", master=frm2, width=4, padx=10)
butRunSeq3 = tk.Button(text="Run", master=frm2, width=4, padx=10)

# Layout of General Inputs frame (frm1)
lblList[19].grid(row=0, column=0, columnspan=3) # Label General & Drive Settings
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
lblList[18].grid(row=2, column=11)

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
inpList[0][8].grid(row=3, column=11)

lblList[13].grid(row=4, column=1) # Channel 2 label
inpList[1][0].grid(row=4, column=2)
inpList[1][1].grid(row=4, column=3)
inpList[1][2].grid(row=4, column=4)
inpList[1][3].grid(row=4, column=5)
inpList[1][4].grid(row=4, column=6)
inpList[1][5].grid(row=4, column=7)
inpList[1][6].grid(row=4, column=8)
butMov2.grid(row=4, column=9)
butStp2.grid(row=4, column=10)
inpList[1][8].grid(row=4, column=11)

lblList[14].grid(row=5, column=1) # Channel 3 label
inpList[2][0].grid(row=5, column=2)
inpList[2][1].grid(row=5, column=3)
inpList[2][2].grid(row=5, column=4)
inpList[2][3].grid(row=5, column=5)
inpList[2][4].grid(row=5, column=6)
inpList[2][5].grid(row=5, column=7)
inpList[2][6].grid(row=5, column=8)
butMov3.grid(row=5, column=9)
butStp3.grid(row=5, column=10)
inpList[2][8].grid(row=5, column=11)

# Layout of Feedback Input frame (frm2)
lblList[20].grid(row=1, column=0, columnspan=2) # Label Feedback Settings
rsmList[0][0].grid(row=2, column=0) # RLS Readout Enable label
inpList[3][5].grid(row=2, column=1) # RLS Readout Enable
rsmList[0][1].grid(row=2, column=2)
rsmList[0][2].grid(row=2, column=3)
rsmList[0][3].grid(row=2, column=4)
rsmList[0][4].grid(row=2, column=5)

rsmList[1][0].grid(row=3, column=1) 
rsmList[1][1].grid(row=3, column=2)
rsmList[1][2].grid(row=3, column=3)
rsmList[1][3].grid(row=3, column=4)
rsmList[1][4].grid(row=3, column=5)

rsmList[2][0].grid(row=4, column=1) 
rsmList[2][1].grid(row=4, column=2)
rsmList[2][2].grid(row=4, column=3)
rsmList[2][3].grid(row=4, column=4)
rsmList[2][4].grid(row=4, column=5)

rsmList[3][0].grid(row=5, column=1) 
rsmList[3][1].grid(row=5, column=2)
rsmList[3][2].grid(row=5, column=3)
rsmList[3][3].grid(row=5, column=4)
rsmList[3][4].grid(row=5, column=5)

butMir1.grid(row=3, column=6)
butMar1.grid(row=3, column=7)
butMir2.grid(row=4, column=6)
butMar2.grid(row=4, column=7)
butMir3.grid(row=5, column=6)
butMar3.grid(row=5, column=7)
butCenter1.grid(row=3, column=8)
butCenter2.grid(row=4, column=8)
butCenter3.grid(row=5, column=8)
butMis1.grid(row=3, column=9)
butMas1.grid(row=3, column=10)
butMis2.grid(row=4, column=9)
butMas2.grid(row=4, column=10)
butMis3.grid(row=5, column=9)
butMas3.grid(row=5, column=10)
butMmr1.grid(row=3, column=11)
butMmr2.grid(row=4, column=11)
butMmr3.grid(row=5, column=11)
inpList[0][7].grid(row=3, column=12)
inpList[1][7].grid(row=4, column=12)
inpList[2][7].grid(row=5, column=12)
butRss.grid(row=6, column=10, columnspan=2)

seqList[0][1].grid(row=2, column=13)
seqList[0][2].grid(row=2, column=14)
seqList[0][3].grid(row=2, column=15)
seqList[0][4].grid(row=2, column=16)
seqList[0][5].grid(row=2, column=17)
seqList[1][1].grid(row=3, column=13)
seqList[2][1].grid(row=4, column=13)
seqList[3][1].grid(row=5, column=13)
seqList[1][2].grid(row=3, column=14)
seqList[2][2].grid(row=4, column=14)
seqList[3][2].grid(row=5, column=14)
seqList[1][3].grid(row=3, column=15)
seqList[2][3].grid(row=4, column=15)
seqList[3][3].grid(row=5, column=15)
seqList[1][4].grid(row=3, column=16)
seqList[2][4].grid(row=4, column=16)
seqList[3][4].grid(row=5, column=16)
seqList[1][5].grid(row=3, column=17)
seqList[2][5].grid(row=4, column=17)
seqList[3][5].grid(row=5, column=17)
butRunSeq1.grid(row=3, column=18)
butRunSeq2.grid(row=4, column=18)
butRunSeq3.grid(row=5, column=18)
seqList[0][6].grid(row=2, column=19, columnspan=2) # Label Sequence Settings
seqList[1][6].grid(row=3, column=19)
seqList[2][6].grid(row=4, column=19)
seqList[3][6].grid(row=5, column=19)
lblList[21].grid(row=1, column=13, columnspan=2) # Label Sequence Settings

     
# Layout of Command History frame (frm4)
v = tk.Scrollbar(frm4, orient='vertical')
v.pack(side=tk.RIGHT, fill='y')     
lblList[16].pack(side=tk.TOP)
txtRespFont = ("Courier New", 12)
txtResp = tk.Text(height = 15, width=140, master=frm4, yscrollcommand=v.set)
txtResp.tag_configure('e', foreground='red')
txtResp.tag_configure('s', foreground='blue')
txtResp.tag_configure('r', foreground='green')
txtResp.configure(font = txtRespFont)
txtResp.pack(side=tk.TOP)

# Setup Info and Checks buttons
butVer = tk.Button(text="Info", master=frm5, padx=10)
butGfs = tk.Button(text="Driver Check", master=frm5, padx=10)
butScm = tk.Button(text="Positioner Check", master=frm5, padx=10)
butTxtRespClear = tk.Button(text="Clear command history", master=frm5, padx=10)
butWinDevMan = tk.Button(text="Device Manager (Windows)", master=frm5)
butInfo = tk.Button(text=">> How To Use <<", master=frm5, width=12, padx=10)
butStages = tk.Button(text="Stage Types", master=frm5, width=12, padx=10)

# Layout of Info and Checks buttons frame (frm5)
butTxtRespClear.grid(row=1, column=1, padx=10)  
frm5.grid_columnconfigure(2, minsize=20)
butVer.grid(row=1, column=3, sticky=tk.W)
butGfs.grid(row=1, column=4, sticky=tk.W)
butScm.grid(row=1, column=5, sticky=tk.W)
butStages.grid(row=1, column=6)
frm5.grid_columnconfigure(7, minsize=20)
butWinDevMan.grid(row=1, column=8, sticky=tk.W)
frm5.grid_columnconfigure(9, minsize=40)
butInfo.grid(row=1, column=10)   # Help Button


# Layout of Disclaimer frame (frm6)
lblList[17].grid(row=1, column=1) 

# Functions
def butVer_handle_click(event):
   butVerThrd=thrd.Thread(target=butVer_handle_thread, daemon=True)
   butVerThrd.start()
    
def butVer_handle_thread():  
   cmdVer = '/VER'
   cmdFiv = 'FIV '
   txtResp.insert('end', ('--> Get firmware version information\n'))
   if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdVer + '\n'), 's') 
   while not rsmReadQueue.empty(): 
       time.sleep(0.1)
   parList[3][6].set(True)    
   try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:   
           response = usbVcp.WriteRead(cmdVer, 1)
           txtResp.insert('end', ('<-- CPSC1 version: ' + response + '\n'), 'r')
           for x in range(6):
               if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdFiv + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdFiv + str(x+1), 1)
               txtResp.insert('end', ('<-- Slot ' + str(x+1) + ': ' + response + '\n'), 'r')            
               txtResp.see("end")
   except IOError:
       txtResp.insert('end', errConnect, 'e')      
   parList[3][6].set(False)

def butStages_handle_click(event):
   butStagesThrd=thrd.Thread(target=butStages_handle_thread, daemon=True)
   butStagesThrd.start()
    
def butStages_handle_thread():  
   cmdStages = '/STAGES'
   txtResp.insert('end', ('--> Get a list of all available Stage Type values\n'))
   if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdStages + '\n'), 's') 
   while not rsmReadQueue.empty(): 
       time.sleep(0.1)
   parList[3][6].set(True)    
   try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:   
           response = usbVcp.WriteRead(cmdStages, 1)
           txtResp.insert('end', ('<-- Stage Type values: ' + response + '\n'), 'r')
   except IOError:
       txtResp.insert('end', errConnect, 'e')      
   parList[3][6].set(False)

def butGfs_handle_click(event):
    butGfsThrd=thrd.Thread(target=butGfs_handle_thread, daemon=True)
    butGfsThrd.start()
     
def butGfs_handle_thread():  
    cmdGfs = 'GFS '
    txtResp.insert('end', ('--> Get CADM2 failsafe state\n'))
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:    
           for x in range(6):
               if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdGfs + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdGfs + str(x+1), 1)
               txtResp.insert('end', ('<-- Slot ' + str(x+1) + ': ' + response + '\n'), 'r')
               txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')        
    parList[3][6].set(False)

def butScm_handle_click(event):
    butScmThrd=thrd.Thread(target=butScm_handle_thread, daemon=True)
    butScmThrd.start()

def butScm_handle_thread():  
    cmdScm = 'SCM '
    txtResp.insert('end', ('--> Check piezo in positioners (will take some time) ... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp: 
           for x in range(6):
               if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdScm + str(x+1) + '\n'), 's')
               response = usbVcp.WriteRead(cmdScm + str(x+1), 1)
               txtResp.insert('end', ('<-- Positioner on slot ' + str(x+1) + ': ' + response + '[nF]\n'), 'r')
               txtResp.see("end")
    except IOError:
       txtResp.insert('end', errConnect, 'e')     
    parList[3][6].set(False)
       
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
    txtResp.insert('end', ('--> Stop movement\n'))
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdStp + '\n'), 's')
    while not rsmReadQueue.empty(): 
        time.sleep(0.1) 
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:          
           response = usbVcp.WriteRead(cmdStp, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
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
    cmdMov = 'MOV ' + str(parList[channel-1][0].get()) + ' ' + str(parList[channel-1][1].get()) + ' ' + str(parList[channel-1][2].get()) + ' ' + str(parList[channel-1][3].get()) + ' ' + str(inpList[channel-1][4].get()) + ' ' + str(parList[3][1].get()) + ' ' + str(parList[channel-1][5].get()) + ' ' + str(parList[channel-1][6].get())
    txtResp.insert('end', ('--> Start movement ...\n'))
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdMov + '\n'), 's')
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMov, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
       txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')    
    parList[3][6].set(False)

def butMir1_handle_click(event):
    butMirThrd=thrd.Thread(target=butMir_handle_thread, args=([1]), daemon=True)
    butMirThrd.start()

def butMir2_handle_click(event):
    butMirThrd=thrd.Thread(target=butMir_handle_thread, args=([2]), daemon=True)
    butMirThrd.start()

def butMir3_handle_click(event):
    butMirThrd=thrd.Thread(target=butMir_handle_thread, args=([3]), daemon=True)
    butMirThrd.start()

def butMir_handle_thread(channel):  
    cmdMir = 'MIR ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) + ' ' + str(parList[channel-1][7].get())
    txtResp.insert('end', ('--> Get current MIR value for CH' + str(channel) + '... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdMir + '\n'), 's')
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMir, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           rsmList[channel][2].config(text=response)
       txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)
    parList[3][5].set(True)

def butMis1_handle_click(event):
    butMisThrd=thrd.Thread(target=butMis_handle_thread, args=([1]), daemon=True)
    butMisThrd.start()

def butMis2_handle_click(event):
    butMisThrd=thrd.Thread(target=butMis_handle_thread, args=([2]), daemon=True)
    butMisThrd.start()

def butMis3_handle_click(event):
    butMisThrd=thrd.Thread(target=butMis_handle_thread, args=([3]), daemon=True)
    butMisThrd.start()

def butMis_handle_thread(channel):  
    cmdMis = 'MIS ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) 
    txtResp.insert('end', ('--> Set current MIS value for CH' + str(channel) + ' ... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdMis + '\n'), 's')
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMis, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)
    parList[3][5].set(True)

def butMar1_handle_click(event):
    butMarThrd=thrd.Thread(target=butMar_handle_thread, args=([1]), daemon=True)
    butMarThrd.start()

def butMar2_handle_click(event):
    butMarThrd=thrd.Thread(target=butMar_handle_thread, args=([2]), daemon=True)
    butMarThrd.start()

def butMar3_handle_click(event):
    butMarThrd=thrd.Thread(target=butMar_handle_thread, args=([3]), daemon=True)
    butMarThrd.start()

def butMar_handle_thread(channel):  
    cmdMar = 'MAR ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) + ' ' + str(parList[channel-1][7].get())
    txtResp.insert('end', ('--> Get current MAR value for CH' + str(channel) + ' ... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdMar + '\n'), 's')
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMar, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           rsmList[channel][3].config(text=response)
       txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)
    parList[3][5].set(True)

def butMas1_handle_click(event):
    butMasThrd=thrd.Thread(target=butMas_handle_thread, args=([1]), daemon=True)
    butMasThrd.start()

def butMas2_handle_click(event):
    butMasThrd=thrd.Thread(target=butMas_handle_thread, args=([2]), daemon=True)
    butMasThrd.start()

def butMas3_handle_click(event):
    butMasThrd=thrd.Thread(target=butMas_handle_thread, args=([3]), daemon=True)
    butMasThrd.start()

def butMas_handle_thread(channel):  
    cmdMas = 'MAS ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) 
    txtResp.insert('end', ('--> Set current MAS value for CH' + str(channel) + '... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdMas + '\n'), 's')
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMas, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)
    parList[3][5].set(True)

def butMmr1_handle_click(event):
    butMmrThrd=thrd.Thread(target=butMmr_handle_thread, args=([1]), daemon=True)
    butMmrThrd.start()

def butMmr2_handle_click(event):
    butMmrThrd=thrd.Thread(target=butMmr_handle_thread, args=([2]), daemon=True)
    butMmrThrd.start()

def butMmr3_handle_click(event):
    butMmrThrd=thrd.Thread(target=butMmr_handle_thread, args=([3]), daemon=True)
    butMmrThrd.start()

def butMmr_handle_thread(channel):  
    cmdMmr = 'MMR ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) 
    txtResp.insert('end', ('--> Reset MIR and MAR values for CH' + str(channel) + '... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdMmr + '\n'), 's')
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdMmr, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)
    parList[3][5].set(True)

def butRss_handle_click(event):
    butRssThrd=thrd.Thread(target=butRss_handle_thread, daemon=True)
    butRssThrd.start()

def butRss_handle_thread():  
    cmdRss = 'RSS ' + str(parList[3][2].get())
    txtResp.insert('end', ('--> Store current MIR and MAR values for all Channels ... \n'))
    parList[3][5].set(False)
    while not rsmReadQueue.empty(): 
        time.sleep(0.1)
    if (parList[3][3].get()): txtResp.insert('end', ('--> ' + cmdRss + '\n'), 's')
    parList[3][6].set(True)
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           response = usbVcp.WriteRead(cmdRss, 1)
           txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
           txtResp.see("end")    
    except IOError:
        txtResp.insert('end', errConnect, 'e')  
    parList[3][6].set(False)    
    parList[3][5].set(True)
        
def butCenter1_handle_click(event):
    butCenterThrd=thrd.Thread(target=butCenter_handle_thread, args=([1]), daemon=True)
    butCenterThrd.start()

def butCenter2_handle_click(event):
    butCenterThrd=thrd.Thread(target=butCenter_handle_thread, args=([2]), daemon=True)
    butCenterThrd.start()

def butCenter3_handle_click(event):
    butCenterThrd=thrd.Thread(target=butCenter_handle_thread, args=([3]), daemon=True)
    butCenterThrd.start()

def butCenter_handle_thread(channel):  
    mirVal = float(rsmList[channel][2].cget("text"))
    marVal = float(rsmList[channel][3].cget("text"))
    centerPos = round(((marVal-mirVal)/2)-marVal,6)
    rsmList[channel][4].config(text=str(centerPos))          
        
def rsmRead_thread():
    while(True):
        if (parList[3][5].get()) and not(parList[3][6].get()):
            rsmReadQueue.put(1)
            nearCenter = 0.0005 # 0.0005m = 0.5mm
            cmdPgva = 'PGVA ' + str(parList[3][2].get()) + ' ' + str(parList[0][7].get()) + ' ' + str(parList[1][7].get()) + ' ' + str(parList[2][7].get())
            try:
               with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
                   response = usbVcp.WriteRead(cmdPgva, 1)
                   #txtResp.insert('end', ('<-- ' + response + '\n'), 'r')
                   responseSplit = response.split(",")
                   if float(responseSplit[0]) < float(rsmList[1][2].cget("text")) or float(responseSplit[0]) > float(rsmList[1][3].cget("text")) :
                       rsmList[1][1].config(text=responseSplit[0], fg='red')   
                   elif float(responseSplit[0]) > float(rsmList[1][4].cget("text"))-nearCenter and float(responseSplit[0]) < float(rsmList[1][4].cget("text"))+nearCenter :
                       rsmList[1][1].config(text=responseSplit[0], fg='blue')  
                   else: 
                       rsmList[1][1].config(text=responseSplit[0], fg='black')  
                   if float(responseSplit[1]) < float(rsmList[2][2].cget("text")) or float(responseSplit[1]) > float(rsmList[2][3].cget("text")) :
                       rsmList[2][1].config(text=responseSplit[1], fg='red')   
                   elif float(responseSplit[1]) > float(rsmList[2][4].cget("text"))-nearCenter and float(responseSplit[1]) < float(rsmList[2][4].cget("text"))+nearCenter :
                       rsmList[2][1].config(text=responseSplit[1], fg='blue')  
                   else: 
                       rsmList[2][1].config(text=responseSplit[1], fg='black')
                   if float(responseSplit[2]) < float(rsmList[3][2].cget("text")) or float(responseSplit[2]) > float(rsmList[3][3].cget("text")) :
                       rsmList[3][1].config(text=responseSplit[2], fg='red')   
                   elif float(responseSplit[2]) > float(rsmList[3][4].cget("text"))-nearCenter and float(responseSplit[2]) < float(rsmList[3][4].cget("text"))+nearCenter :
                       rsmList[3][1].config(text=responseSplit[2], fg='blue')  
                   else: 
                       rsmList[3][1].config(text=responseSplit[2], fg='black')    
                                                             
            except IOError:
                #txtResp.insert('end', 'communication lost', 'e')  
                pass      
            
            rsmReadQueue.get(1)
        time.sleep(0.1)   

def txtResp_clear_click(event):
    txtResp.delete("1.0", "end")

def butWinDevMan_clear_click(event):
    sp.call("control /name Microsoft.DeviceManager")
    
def butInfo_click(event):
    # Display some initial tips and hints
    txtResp.insert('end', ('1) GENERAL SETTINGS\n'
                           'A) Set COM port, Baudrate and CADM and RSM addressing first.\n\n'                      
                           '2) DRIVE SETTINGS\n'
                           'A) Make sure to set [Direction], [Freq], [Rss], [# Steps], [Temperature] and correct [Stage Type] before clicking the Move buttons. [Drive Factor] can be left at default value (1.0).\n'
                           'B) Use [Move] and [Stop] buttons to drive a positioner (BaseDrive mode of operation).\n'
                           'C) [Positioner ID] is currently only used to name log data files generated with the "SEQUENCE RUN" function.\n\n'                     
                           '3) FEEDBACK SETTINGS\n'
                           'A) Check [RSM Polling] to start polling the RLS values continuously.\n'
                           'B) Set the minimum travel position [MIS] and maximum travel position [MAS] (use [MIR] and [MAR] to get current stored values).\n'
                           'C) Make sure to select the appropriate [Stage Type] before using the [MIS] and [MAS] functions. Click [Store Settings] to save current values to controller memory.\n\n'
                           '4) SEQUENCE RUN\n'
                           'A) Make sure to set the Minimum position [MIS] and maximum position [MAS] and correct [Stage Type] before clicking the Run buttons.\n'
                           'B) Clicking the "Run" button will move a positioner between Pos A and Pos B for "#Runs" number of times.\n'
                           'C) "Timer" shows the elapsed time between sequence runs. "Seq#" is only used to name log data files.\n'
                           'D) "Timeout" is used an escape when positioners do not reach PosA or PosB within set time.\n'
                           'E) After completing a sequence a .PNG plot will be stored as well as a .CSV file containing the raw position data.\n\n'))
                         
def doNothing(event):
    time.sleep(1)

def butRunSeq1_handle_click(event): 
    butRunSeqThrd=thrd.Thread(target=butRunSeq_handle_thread, args=([1]), daemon=True)
    butRunSeqThrd.start()

def butRunSeq2_handle_click(event): 
    butRunSeqThrd=thrd.Thread(target=butRunSeq_handle_thread, args=([2]), daemon=True)
    butRunSeqThrd.start()

def butRunSeq3_handle_click(event): 
    butRunSeqThrd=thrd.Thread(target=butRunSeq_handle_thread, args=([3]), daemon=True)
    butRunSeqThrd.start()
    
def butRunSeq_handle_thread(channel):    
    parList[3][5].set(False)        # Disable continuous RSM readout first
    while not rsmReadQueue.empty(): # Wait until queue is empty
        time.sleep(0.1)
    parList[3][6].set(True)         # Command 'lock' on
    
    # Set some default values for variables used in this routine
    startTime = 0
    passedTime = 0
    logTimeA = []
    logTimeB = []
    logRlsA = []
    logRlsB = []
    logElapA = []
    logElapB = []
    pollDelay = 0.1
    sequenceDelay = 2.0
    
    # Set up CPSC commands that are used in this routine
    cmdMvA = 'MOV ' + str(parList[channel-1][0].get()) + ' 0 ' + str(parList[channel-1][2].get()) + ' ' + str(parList[channel-1][3].get()) + ' 0 ' + str(parList[3][1].get()) + ' ' + str(parList[channel-1][5].get()) + ' ' + str(parList[channel-1][6].get())
    cmdMvB = 'MOV ' + str(parList[channel-1][0].get()) + ' 1 ' + str(parList[channel-1][2].get()) + ' ' + str(parList[channel-1][3].get()) + ' 0 ' + str(parList[3][1].get()) + ' ' + str(parList[channel-1][5].get()) + ' ' + str(parList[channel-1][6].get())
    cmdStp = 'STP ' + str(parList[channel-1][0].get())
    cmdMir = 'MIR ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) + ' ' + str(parList[channel-1][7].get())
    cmdMar = 'MAR ' + str(parList[3][2].get()) + ' ' + str(parList[channel-1][0].get()) + ' ' + str(parList[channel-1][7].get())
    cmdPgva = 'PGVA ' + str(parList[3][2].get()) + ' ' + str(parList[0][7].get()) + ' ' + str(parList[1][7].get()) + ' ' + str(parList[2][7].get())

    # Initial warning message
    txtResp.insert('end', ('==> Start Move Sequence. Please be patient! There is currently no abort option.\n')) 
    txtResp.see("end") 
    time.sleep(sequenceDelay)
 
    try:
       with CpscSerial.CpscSerialInterface(('COM' + str(parList[3][0].get())), str(inpList[3][4].get())) as usbVcp:        
           txtResp.insert('end', ('==> Get current MIR value for CH' + str(channel) + '... \n'))
           if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMir + '\n'), 's')   
           response = usbVcp.WriteRead(cmdMir, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           rsmList[channel][2].config(text=response)
           txtResp.see("end")    
        
           txtResp.insert('end', ('==> Get current MAR value for CH' + str(channel) + ' ... \n'))
           if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMar + '\n'), 's')
           response = usbVcp.WriteRead(cmdMar, 1)
           txtResp.insert('end', ('<== ' + response + '\n'), 'r')
           rsmList[channel][3].config(text=response)
           txtResp.see("end")    
           
           # Only run sequence if Pos A and Pos B do not exceed MIR/MAR values
           if float(seqList[channel][1].get()) > float(rsmList[channel][2].cget("text")) and float(seqList[channel][2].get()) < float(rsmList[channel][3].cget("text")):       
                txtResp.insert('end', ('==> Get current RLS values for all channels ... \n'))
                response = usbVcp.WriteRead(cmdPgva, 1)
                responseSplit = response.split(",")
                rsmList[1][1].config(text=responseSplit[0], fg='black')  
                rsmList[2][1].config(text=responseSplit[1], fg='black')
                rsmList[3][1].config(text=responseSplit[2], fg='black')   
                txtResp.see("end")                                                        
        
                time.sleep(sequenceDelay)
        
                # Move to initial position ("POS A")
                if float(responseSplit[channel-1]) > float(seqList[channel][1].get()): # If current position is positive to Pos A, move in DIR=0 direction towards Pos A
                    txtResp.insert('end', ('==> Start moving in DIR=0 until ' + seqList[channel][1].get() + '[m] reached ...\n'))
                    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMvA + '\n'), 's')
                    startTime = time.time()  
                    response = usbVcp.WriteRead(cmdMvA, 1)
                    txtResp.insert('end', ('<== ' + response + '\n'), 'r') 
                    txtResp.see("end")  
                    time.sleep(pollDelay) 
                    while float(responseSplit[channel-1]) >= float(seqList[channel][1].get()) and not passedTime > int(seqList[channel][5].get()): # Keep polling RLS value until Pos A reached OR timeout has occurred
                        response = usbVcp.WriteRead(cmdPgva, 1)
                        responseSplit = response.split(",")
                        rsmList[1][1].config(text=responseSplit[0], fg='black')  
                        rsmList[2][1].config(text=responseSplit[1], fg='black')
                        rsmList[3][1].config(text=responseSplit[2], fg='black')
                        passedTime = time.time() - startTime                                                         
                        seqList[channel][6].config(text=str(round(passedTime, 1)), fg='black') 
                        time.sleep(pollDelay)
                                                                    
                    txtResp.insert('end', ('==> Stop moving\n'))
                    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
                    response = usbVcp.WriteRead(cmdStp, 1)
                    txtResp.insert('end', ('<== ' + response + '\n'), 'r')    
                    txtResp.see("end")                 
                    
                else: # If current position is negative to Pos A, move in DIR=1 direction towards Pos A
                    txtResp.insert('end', ('==> Start moving in DIR=1 until ' + seqList[channel][1].get() + '[m] reached ...\n'))
                    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMvB + '\n'), 's')
                    startTime = time.time()  
                    response = usbVcp.WriteRead(cmdMvB, 1)
                    txtResp.insert('end', ('<== ' + response + '\n'), 'r')  
                    txtResp.see("end")  
                    time.sleep(pollDelay)
                    while float(responseSplit[channel-1]) >= float(seqList[channel][1].get()) and not passedTime > int(seqList[channel][5].get()): # Keep polling RLS value until Pos A reached OR timeout has occurred
                        response = usbVcp.WriteRead(cmdPgva, 1)
                        responseSplit = response.split(",")
                        rsmList[1][1].config(text=responseSplit[0], fg='black')  
                        rsmList[2][1].config(text=responseSplit[1], fg='black')
                        rsmList[3][1].config(text=responseSplit[2], fg='black') 
                        passedTime = time.time() - startTime   
                        seqList[channel][6].config(text=str(round(passedTime, 1)), fg='black')                                                       
                        time.sleep(pollDelay)
             
                    txtResp.insert('end', ('==> Stop moving\n'))
                    if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
                    response = usbVcp.WriteRead(cmdStp, 1)
                    txtResp.insert('end', ('<== ' + response + '\n'), 'r')   
                    txtResp.see("end")  
        
                if passedTime > int(seqList[channel][5].get()):
                    txtResp.insert('end', ('==> Timeout occured! Starting Position A has not been reached.\n'), 'e')
                    txtResp.see("end")
                    time.sleep(sequenceDelay)              
                else:
                    # Finally start the actual sequence run
                    txtResp.insert('end', ('==> Position A reached\n'))
                    txtResp.see("end")
                    time.sleep(sequenceDelay)
                    for i in range(0, int(seqList[channel][3].get())):
                        # Move towards Pos B
                        txtResp.insert('end', ('==> Run #' + str(i+1) + '. Move to ' + seqList[channel][2].get() + '[m] ...\n'))
                        if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMvB + '\n'), 's')
                        passedTime = 0
                        startTime = time.time()       
                        response = usbVcp.WriteRead(cmdMvB, 1)
                        txtResp.insert('end', ('<== ' + response + '\n'), 'r')  
                        txtResp.see("end")  
                        response = usbVcp.WriteRead(cmdPgva, 1)
                        responseSplit = response.split(",")
                        rsmList[1][1].config(text=responseSplit[0], fg='black')  
                        rsmList[2][1].config(text=responseSplit[1], fg='black')
                        rsmList[3][1].config(text=responseSplit[2], fg='black')     
                        while float(responseSplit[channel-1]) <= float(seqList[channel][2].get()) and not round((passedTime),1) > int(seqList[channel][5].get()): # Keep polling RLS value until Pos B reached OR timeout has occurred
                            response = usbVcp.WriteRead(cmdPgva, 1)
                            responseSplit = response.split(",")
                            rsmList[1][1].config(text=responseSplit[0], fg='black')  
                            rsmList[2][1].config(text=responseSplit[1], fg='black')
                            rsmList[3][1].config(text=responseSplit[2], fg='black')                                
                            logRlsB.append(float(responseSplit[channel-1]))   
                            passedTime = time.time() - startTime
                            logTimeB.append(passedTime)
                            seqList[channel][6].config(text=str(round(passedTime, 1)), fg='black') 
                            time.sleep(pollDelay)
                            
                        txtResp.insert('end', ('==> Stop moving\n'))
                        if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
                        response = usbVcp.WriteRead(cmdStp, 1)
                        txtResp.insert('end', ('<== ' + response + '\n'), 'r')   
                        txtResp.see("end")  
                        
                        if passedTime > int(seqList[channel][5].get()):
                            txtResp.insert('end', ('==> Timeout occured! Position B has not been reached at Run #' + str(i+1) + '\n'), 'e')
                            txtResp.see("end")
                        
                        logElapB.append(passedTime)
                        time.sleep(sequenceDelay) 
                        
                        # Move towards Pos A
                        txtResp.insert('end', ('==> Run #' + str(i+1) + '. Move to ' + seqList[channel][1].get() + '[m] ...\n'))
                        if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdMvA + '\n'), 's')
                        passedTime = 0
                        startTime = time.time()   
                        response = usbVcp.WriteRead(cmdMvA, 1)
                        txtResp.insert('end', ('<== ' + response + '\n'), 'r')  
                        txtResp.see("end")  
                        response = usbVcp.WriteRead(cmdPgva, 1)
                        responseSplit = response.split(",")
                        rsmList[1][1].config(text=responseSplit[0], fg='black')  
                        rsmList[2][1].config(text=responseSplit[1], fg='black')
                        rsmList[3][1].config(text=responseSplit[2], fg='black')     
                        while float(responseSplit[channel-1]) >= float(seqList[channel][1].get()) and not round((passedTime),1) > int(seqList[channel][5].get()): # Keep polling RLS value until Pos A reached OR timeout has occurred
                            response = usbVcp.WriteRead(cmdPgva, 1)
                            responseSplit = response.split(",")
                            rsmList[1][1].config(text=responseSplit[0], fg='black')  
                            rsmList[2][1].config(text=responseSplit[1], fg='black')
                            rsmList[3][1].config(text=responseSplit[2], fg='black')                                                         
                            logRlsA.append(float(responseSplit[channel-1]))   
                            passedTime = time.time() - startTime
                            logTimeA.append(passedTime)
                            seqList[channel][6].config(text=str(round(passedTime, 1)), fg='black') 
                            time.sleep(pollDelay)
                 
                        txtResp.insert('end', ('==> Stop moving\n'))
                        if (parList[3][3].get()): txtResp.insert('end', ('==> ' + cmdStp + '\n'), 's')
                        response = usbVcp.WriteRead(cmdStp, 1)
                        txtResp.insert('end', ('<== ' + response + '\n'), 'r')  
                        txtResp.see("end")  
                        
                        if passedTime > int(seqList[channel][5].get()):
                            txtResp.insert('end', ('==> Timeout occured! Position A has not been reached at Run #' + str(i+1) + '\n'), 'e')
                            txtResp.see("end")
                        
                        logElapA.append(passedTime)
                        time.sleep(sequenceDelay)     
                                 
                    txtResp.insert('end', ('==> Movement Sequence completed\n'))
                    txtResp.see("end")    
        
           else:
               txtResp.insert('end', ('==> Sequence run not started: Pos A or Pos B outside of Minimum (MIR) or Maximum (MAR) travel range!\n'), 'e') 
               txtResp.see("end")         
           
            
       a = sum(logElapA) / len(logElapA)
       b = sum(logElapB) / len(logElapB)
       c = (float(seqList[channel][2].get())-float(seqList[channel][1].get()))*1000
       d = round(c / a, 2)
       e = round(c / b, 2)  
       
       #Consolde debug info
       print('[AVRG Time B->A]: ' + str(a) + '[s]')
       print('[AVRG Time A->B]: ' + str(b) + '[s]')
       print('[Distance A<->B]: ' + str(c) + '[mm]')
       print('Speed B->A: ' + str(d) + '[mm/s]')
       print('Speed A->B: ' + str(e) + '[mm/s]')
       
       txtResp.insert('end', ('==> Create plot and store data to file ... \n'))
       fig, (ax1, ax2) = plt.subplots(2)
       fig.suptitle('Sequence #' + str(seqList[channel][4].get()) + ' of ' + str(parList[channel-1][5].get()) + ' #' + str(inpList[channel-1][8].get()) +
                    '\nTest Params: [FREQ]:' + str(parList[channel-1][2].get()) + ' [RSS]:' + str(parList[channel-1][3].get()) + ' [TEMP]:' + str(parList[3][1].get()) + 
                    ' [DF]:' + str(parList[channel-1][6].get()) + ' [#Runs]:' + str(seqList[channel][3].get()) + 
                    '\n[MIR]:' + str(rsmList[channel][2].cget("text")) + ' [MAR]:' + str(rsmList[channel][3].cget("text")) + 
                    ' [POSA]:' + str(seqList[channel][1].get()) + ' [POSB]:' + str(seqList[channel][2].get()) +
                    '\n[AVRG Speed B->A]: ' + str(d) + '[mm/s]' + ' [AVRG Speed A->B]: ' + str(e) + '[mm/s]', fontsize=7)  # CALCULATION NOT CORRECT!!
       
       ax1.set_xlabel('Time [s]')
       ax1.set_ylabel('RLS Position [m]', color='b')
       ax1.plot(logTimeB, logRlsB, color='b', linestyle='None', marker='.', markersize=3) 
       ax1.tick_params(axis='y', labelcolor='b')
       ax1.grid(which='major', linestyle='--', linewidth=0.2)
       ax1.grid(which='minor', linestyle='-', linewidth=0.1)
       ax1.minorticks_on()
       ax2.set_xlabel('Time [s]')
       ax2.set_ylabel('RLS Position [m]', color='g')
       ax2.plot(logTimeA, logRlsA, color='g', linestyle='None', marker='.', markersize=3) 
       ax2.tick_params(axis='y', labelcolor='g')
       ax2.grid(which='major', linestyle='--', linewidth=0.2)
       ax2.grid(which='minor', linestyle='-', linewidth=0.1)
       ax2.minorticks_on()
                
       plt.figtext(0.01, 0.01, str(sys.argv[0]), fontsize=5)
       fig.tight_layout()
         
       plt.savefig(str(parList[channel-1][5].get()) + '_' + str(inpList[channel-1][8].get()) + '-' + str(seqList[channel][4].get()) + '.png', dpi=300)
       plt.show()
       #plt.close()
         
       logData = list(zip(logTimeB, logRlsB, logTimeA, logRlsA))
        
       with open(str(parList[channel-1][5].get()) + '_' + str(inpList[channel-1][8].get()) + '-' + str(seqList[channel][4].get()) + '.csv', 'w', newline="") as x:
           csv.writer(x, delimiter=";", dialect='excel').writerows(logData)
       x.close()        

       txtResp.insert('end', ('==> Sequence Run function finished!\n'))
       txtResp.see("end")       
        
    except IOError:
        txtResp.insert('end', '==> Communication lost', 'e') 
        pass        

    parList[3][6].set(False)

# Bind button press (left mouse click) events to functions
butVer.bind("<Button-1>", butVer_handle_click)
butGfs.bind("<Button-1>", butGfs_handle_click)
butScm.bind("<Button-1>", butScm_handle_click)
butMov1.bind("<Button-1>", butMov1_handle_click)
butMov2.bind("<Button-1>", butMov2_handle_click)
butMov3.bind("<Button-1>", butMov3_handle_click)
butStp1.bind("<Button-1>", butStp1_handle_click)
butStp2.bind("<Button-1>", butStp2_handle_click)
butStp3.bind("<Button-1>", butStp3_handle_click)
butMir1.bind("<Button-1>", butMir1_handle_click)
butMir2.bind("<Button-1>", butMir2_handle_click)
butMir3.bind("<Button-1>", butMir3_handle_click)
butMar1.bind("<Button-1>", butMar1_handle_click)
butMar2.bind("<Button-1>", butMar2_handle_click)
butMar3.bind("<Button-1>", butMar3_handle_click)
butCenter1.bind("<Button-1>", butCenter1_handle_click)
butCenter2.bind("<Button-1>", butCenter2_handle_click)
butCenter3.bind("<Button-1>", butCenter3_handle_click)
butMis1.bind("<Button-1>", butMis1_handle_click)
butMis2.bind("<Button-1>", butMis2_handle_click)
butMis3.bind("<Button-1>", butMis3_handle_click)
butMas1.bind("<Button-1>", butMas1_handle_click)
butMas2.bind("<Button-1>", butMas2_handle_click)
butMas3.bind("<Button-1>", butMas3_handle_click)
butMmr1.bind("<Button-1>", butMmr1_handle_click)
butMmr2.bind("<Button-1>", butMmr2_handle_click)
butMmr3.bind("<Button-1>", butMmr3_handle_click)
butRss.bind("<Button-1>", butRss_handle_click)
butRunSeq1.bind("<Button-1>", butRunSeq1_handle_click)
butRunSeq2.bind("<Button-1>", butRunSeq2_handle_click)
butRunSeq3.bind("<Button-1>", butRunSeq3_handle_click)
butTxtRespClear.bind("<Button-1>", txtResp_clear_click)
butWinDevMan.bind("<Button-1>", butWinDevMan_clear_click)
butInfo.bind("<Button-1>", butInfo_click)
butStages.bind("<Button-1>", butStages_handle_click)

butVer.bind("<Double-Button-1>", doNothing)
butGfs.bind("<Double-Button-1>", doNothing)
butScm.bind("<Double-Button-1>", doNothing)
butMov1.bind("<Double-Button-1>", doNothing)
butMov2.bind("<Double-Button-1>", doNothing)
butMov3.bind("<Double-Button-1>", doNothing)
butStp1.bind("<Double-Button-1>", doNothing)
butStp2.bind("<Double-Button-1>", doNothing)
butStp3.bind("<Double-Button-1>", doNothing)
butMir1.bind("<Double-Button-1>", doNothing)
butMir2.bind("<Double-Button-1>", doNothing)
butMir3.bind("<Double-Button-1>", doNothing)
butMar1.bind("<Double-Button-1>", doNothing)
butMar2.bind("<Double-Button-1>", doNothing)
butMar3.bind("<Double-Button-1>", doNothing)
butCenter1.bind("<Double-Button-1>", doNothing)
butCenter2.bind("<Double-Button-1>", doNothing)
butCenter3.bind("<Double-Button-1>", doNothing)
butMis1.bind("<Double-Button-1>", doNothing)
butMis2.bind("<Double-Button-1>", doNothing)
butMis3.bind("<Double-Button-1>", doNothing)
butMas1.bind("<Double-Button-1>", doNothing)
butMas2.bind("<Double-Button-1>", doNothing)
butMas3.bind("<Double-Button-1>", doNothing)
butMmr1.bind("<Double-Button-1>", doNothing)
butMmr2.bind("<Double-Button-1>", doNothing)
butMmr3.bind("<Double-Button-1>", doNothing)
butRss.bind("<Double-Button-1>", doNothing)
butRunSeq1.bind("<Double-Button-1>", doNothing)
butRunSeq2.bind("<Double-Button-1>", doNothing)
butRunSeq3.bind("<Double-Button-1>", doNothing)
butTxtRespClear.bind("<Double-Button-1>", doNothing)
butWinDevMan.bind("<Double-Button-1>", doNothing)
butInfo.bind("<Double-Button-1>", doNothing)
butStages.bind("<Double-Button-1>", doNothing)

# Bind key press events to functions  
window.bind("<Insert>", butMov1_handle_click)
window.bind("<Home>", butMov2_handle_click)
window.bind("<Prior>", butMov3_handle_click)
window.bind("<Delete>", butStp1_handle_click)  
window.bind("<End>", butStp2_handle_click)  
window.bind("<Next>", butStp3_handle_click)  

# Configure and start RSM read out thread
rsmReadQueue = Queue()
rsmReadThrd=thrd.Thread(target=rsmRead_thread)
rsmReadThrd.daemon = True
rsmReadThrd.start()

# Main loop (loop until window is closed)
window.mainloop()
