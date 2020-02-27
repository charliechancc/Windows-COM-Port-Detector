import win32com.client
import re

wmi = win32com.client.GetObject("winmgmts:")
counter = 0
for serial in wmi.InstancesOf("Win32_SerialPort"):
	start = serial.Name.find("(") + 1
	end   = serial.Name.find(")")
	print (serial.Name[start: end] + " : " + serial.Description)
	counter += 1

print(counter)
if counter > 0:
	print ()
	print ("Found " + str(counter) + " COM Port", end = '')
	
	if counter > 1:
		print ("s", end = '')
else:
	print ("Cannot detect any COM Port")
	
input()