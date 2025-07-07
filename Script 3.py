# -*- coding: utf-8 -*-
"""
Created on Wed Jul 17 09:53:12 2024

@author: ISTM
"""

import os
from datetime import datetime

folder_path = r"C:\Users\ISTM\OneDrive - Ramboll\Desktop\ISHANT\geotech\Hs2 Curzon\SLS CH for WP07 06Dec2024"

# Use os.scandir() to list files
with os.scandir(folder_path) as entries:
    # Get all files with their paths and modification times
    files = [(entry.name, entry.stat().st_mtime) for entry in entries if entry.is_file()]

# Sort files by their modification time


# Extract sorted file 



sorted_names = [file[0] for file in files]

# Remove the last four characters from each filename (assuming they are file extensions)
modified_names = [name[:-4] for name in sorted_names]

print("Sorted file names (without extensions):")
for name in modified_names:
    print(name)

# Optional: Print file names with modification dates
#print("\nSorted file names with modification dates:")
#for name, mtime in files:
    #print(f"{name}: {datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')}")  
    

import pyautogui

#f=open("Pile group 4.0.,4.2,5.0.,6.0.(2-7) Type 2.rpx",'r')
#print(f.read())





b=0.1 #between steps
c=0.01 #between clicks
#d=1 #drag the window
pause_time=1
pause_time_PC=0.5
pause_time_Pl=0.5
drag_speed_pileresponse=0.5 #1

for d in range(0,16):  #for d in range(len(modified_names)): 63
  
  if d < 39:
      col=270
      p=0
  else:
      col=504
      p=p+1
    
  pyautogui.moveTo(col, 180+21*(d-p), 0.01)
  
 
  pyautogui.click(x=col, y=180+21*(d-p), clicks=2, interval=0.01, button='left') #open pile group
  p=38
  pyautogui.PAUSE = 15
  #15
  

  print(f'opened pile group {modified_names[d]}')


#pyautogui.hotkey('win','up') 
#pyautogui.hotkey('alt','tab')
  pyautogui.click(x=927, y=563, clicks=1, interval=0.5, button='left')#view
  pyautogui.PAUSE = pause_time-0.9
  pyautogui.click(x=947, y=683, clicks=1, interval=0.5, button='left')#view
  pyautogui.PAUSE = pause_time-0.9
  pyautogui.click(x=907, y=523, clicks=1, interval=0.5, button='left')#view
  pyautogui.PAUSE = pause_time-0.9
  pyautogui.click(x=987, y=563, clicks=1, interval=0.5, button='left')#view
  

# Press x to maximize the window
  
  pyautogui.hotkey('win','up')
 # pyautogui.press('f11') 
  #pyautogui.PAUSE = pause_time
  print("maximised")
#pyautogui.click(x=3367, y=102, clicks=1, interval=1, button='left') #maximise 

  #pyautogui.PAUSE = pause_time
#pyautogui.typewrite('Hello world!\n', 0.5)C1
  pyautogui.click(x=282, y=49, clicks=1, interval=c, button='left')#Build
  print("clicked build in repute")
  #pyautogui.PAUSE = pause_time

  pyautogui.click(x=640, y=99, clicks=1, interval=c, button='left') #rerun  all


  pyautogui.PAUSE = 5#10
  print("reran all the analysis")  
  
  pyautogui.click(x=789, y=45, clicks=1, interval=c, button='left')#click workbook
  print("clicked workbook") 
  
  
  
  pyautogui.moveTo(297, 449, 0.01)
  
  pyautogui.dragTo(953, 449, 0.5, button='left')
  
  print("expand the window to the center") 
  #pyautogui.click(x=248, y=122, clicks=1, interval=0.01, button='left')#click workbook
  
  for h in range(1,7):
        
        #print("BEA click arrow") 
          pyautogui.PAUSE = pause_time
          pyautogui.click(x=248, y=122, clicks=1, interval=0.1, button='left')
          pyautogui.PAUSE = pause_time
        #pyautogui.moveTo(248, 143+15*(h-1), 0.1)
          pyautogui.click(x=248, y=143+15*(h-1), clicks=1, interval=0.1, button='left')
          pyautogui.PAUSE = pause_time
  
          pyautogui.click(x=497, y=84, clicks=1, interval=0.1, button='left')
          pyautogui.PAUSE = pause_time
        
          pyautogui.typewrite(f'{modified_names[d]}_PC_LC{h}', interval=0.01) #write name
          print("named the pile cap")
          pyautogui.press('enter')

  for h in range(1,7): #7
          pyautogui.PAUSE = 3
          pyautogui.click(x=248, y=122, clicks=1, interval=0.1, button='left')
          #pyautogui.PAUSE = pause_time_Pl
          #pyautogui.moveTo(248, 143+15*(h-1), 0.1)
          pyautogui.click(x=248, y=143+15*(h-1), clicks=1, interval=0.1, button='left')
          #pyautogui.PAUSE = pause_time_Pl
          
          pyautogui.click(x=489, y=733, clicks=1, interval=c, button='left')#
          #pyautogui.PAUSE = pause_time_Pl+2
          print("click pile response") #x=348
          
          pyautogui.click(x=124, y=296, clicks=1, interval=c, button='right')#
          #pyautogui.PAUSE = pause_time_Pl+2
          print("right click to open the dropdwon to select the feld chooser ")
          
          pyautogui.moveTo(170, 521, 0.01)
          #pyautogui.PAUSE = pause_time_Pl+2
          pyautogui.click(x=170, y=521, clicks=1, interval=c, button='left')#
          #pyautogui.PAUSE = pause_time_Pl+2
          print("select the feld chooser ")
          
          pyautogui.click(x=782, y=464, clicks=1, interval=c, button='left')#
          #pyautogui.PAUSE = pause_time_Pl
          print("click the coloumn ")
          
        
          pyautogui.moveTo(729, 540, 0.1) #200
          pyautogui.PAUSE = 5
          pyautogui.dragTo(177, 286, drag_speed_pileresponse, button='left')
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.moveTo(729, 544, 0.1)
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.dragTo(235, 284, drag_speed_pileresponse, button='left')
          pyautogui.PAUSE = pause_time_Pl
          print("drag fx and fy ")
              
        
          pyautogui.moveTo(729, 563, 0.1)
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.dragTo(419, 286, drag_speed_pileresponse, button='left')
          pyautogui.PAUSE = pause_time_Pl
          
          pyautogui.moveTo(729, 560, 0.1)
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.dragTo(484, 286, drag_speed_pileresponse, button='left') 
          pyautogui.PAUSE = pause_time_Pl
          print("drag My and Mz")
              
          pyautogui.click(x=856, y=697, clicks=3, interval=0.1, button='left')
          pyautogui.PAUSE = pause_time_Pl
          #pyautogui.moveTo(853
    
          pyautogui.moveTo(729, 522, 0.1)
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.dragTo(740, 286, drag_speed_pileresponse, button='left')
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.moveTo(729, 545, 0.1)
          pyautogui.PAUSE = pause_time_Pl
          pyautogui.dragTo(804, 286, drag_speed_pileresponse, button='left')  
          pyautogui.PAUSE = pause_time_Pl
          print("drag sx and sy")
              
              
          pyautogui.click(x=497, y=84, clicks=1, interval=0.1, button='left')
          #pyautogui.PAUSE = pause_time_Pl
          #pyautogui.PAUSE = b
        
        #pyautogui.click(x=2038+2, y=347+2, clicks=1, interval=1, button='left')#select the export
        #pyautogui.PAUSE = b
          pyautogui.typewrite(f'{modified_names[d]}_LC{h}', interval=0.01) #write name
          
          print("named the pile group")
         
          #pyautogui.PAUSE = b
    
          pyautogui.press('enter')
  #pyautogui.PAUSE = b

  pyautogui.click(x=1904, y=23, clicks=1, interval=0.1, button='left') #close file
  pyautogui.PAUSE = pause_time
  pyautogui.click(x=1038, y=609, clicks=1, interval=0.5, button='left')
  pyautogui.PAUSE = pause_time
  print("file not saved and closed")

  



