'''
ForeLook.py brings outlook to the foreground on your desktop if the system has been idle for more than 3 minutes
This module reminds you to check your emails if you have been away. :P
'''

__author__ = 'RRShend- Rishikesh Shendkar'

import win32gui
from ctypes import Structure, windll, c_uint, sizeof, byref
import time
 
def windowEnumerationHandler(hwnd, top_windows):
	top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
 
class LASTINPUTINFO(Structure):
	_fields_ = [
		('cbSize', c_uint),
		('dwTime', c_uint),
	]
 
def get_idle_duration():
	lastInputInfo = LASTINPUTINFO()
	lastInputInfo.cbSize = sizeof(lastInputInfo)
	windll.user32.GetLastInputInfo(byref(lastInputInfo))
	millis = windll.kernel32.GetTickCount() - lastInputInfo.dwTime
	return millis / 1000.0

if __name__ == "__main__":
	GetLastInputInfo = int(get_idle_duration())
	print(GetLastInputInfo)
	if GetLastInputInfo > 180:
		# if GetLastInputInfo is 8 minutes, play a sound
		results = []
		top_windows = []
		win32gui.EnumWindows(windowEnumerationHandler, top_windows)
		for i in top_windows:
			if "outlook" in i[1].lower():
				print(i)
				win32gui.ShowWindow(i[0],3)
				win32gui.SetForegroundWindow(i[0])
				break