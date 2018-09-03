import webbrowser
import pyautogui
import time
import os
import fnmatch
import shutil
#import subprocess

#pyautogui.position()

downloads_dir = 'C:\\Users\\LeonardoMaireneMuniz\\Downloads\\'
brrt_dir = 'D:\\brrt\\data\\'

#open the url below
#subprocess.call([r'C:\\Program Files\\Internet Explorer\\iexplore.exe',
#                 'https://w3.ibm.com/tools/mobile-management/portal/app/getall'])

webbrowser.open('https://w3.ibm.com/tools/mobile-management/portal/app/getall')

#wait 10 sec for page to open
time.sleep(10)
#pyautogui.hotkey('esc')

#point mouse over the e-mail field and clear content
pyautogui.moveTo(443, 383, duration=0.25)
pyautogui.click(443, 383, button='left', duration=0.25)
pyautogui.hotkey('ctrl','a')
time.sleep(1)
pyautogui.hotkey('del')

#wait 1 sec and point mouse over the e-mail field and type hubrpts@us.ibm.com e-mail address
time.sleep(1)
pyautogui.typewrite('hubrpts@us.ibm.com')

#wait 1 sec and press tab to move to pwd field and type in the hubrpt@us.ibm.com password
time.sleep(1)
pyautogui.hotkey('tab')
pyautogui.typewrite(os.environ['MyVar'])

#wait 1 sec and press enter for login
time.sleep(1)
pyautogui.press('enter')

#click on the filter by name field and type in mysa and click enter
#time.sleep(20)
#pyautogui.click(240, 305, button='left')
#pyautogui.typewrite('MySA')
#time.sleep(1)
#pyautogui.press('enter')

#wait 20 sec for page to load and move mouse and click to first sheet to download
time.sleep(30)
pyautogui.moveTo(83,384, duration=0.25)
pyautogui.click(83,384, button='left')


#pyautogui.hotkey('alt','d')
#pyautogui.press('enter')
recent_dowloand_file_name = '' 

time.sleep(30)

for entry in os.listdir(downloads_dir):  
    if fnmatch.fnmatch(entry, "MySA*.xlsx"):
        recent_dowloand_file_name = entry

if recent_dowloand_file_name != '':
    print('moving file... C:\\Users\\LeonardoMaireneMuniz\\Downloads\\' + recent_dowloand_file_name)
    shutil.move('C:\\Users\\LeonardoMaireneMuniz\\Downloads\\' + recent_dowloand_file_name, 'D:\\brrt\\data\\MySA_Android_downloads.xlsx')  
    


#wait 20 sec for page to load and move mouse and click to second sheet to download
time.sleep(5)
pyautogui.moveTo(83,424, duration=0.25)
pyautogui.click(83,424, button='left')

recent_dowloand_file_name = '' 

time.sleep(30)

for entry in os.listdir(downloads_dir):  
    if fnmatch.fnmatch(entry, "MySA*.xlsx"):
        recent_dowloand_file_name = entry
       
if recent_dowloand_file_name != '':
    print('moving file... C:\\Users\\LeonardoMaireneMuniz\\Downloads\\' + recent_dowloand_file_name)
    shutil.move('C:\\Users\\LeonardoMaireneMuniz\\Downloads\\' + recent_dowloand_file_name, 'D:\\brrt\\data\\MySA_Apple_downloads.xlsx')   


#wait 30 sec for second file to download and close web browser
time.sleep(5)
pyautogui.hotkey('alt', 'f4')