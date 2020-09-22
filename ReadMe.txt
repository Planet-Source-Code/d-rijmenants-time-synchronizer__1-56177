
 TIME SYNC - README

 This program synchronizes the computers system time
 with one of the selected time timeservers.

 To execute synchroniation directly without userinterface,
 create a shortcut with the "\now" prompt. When placed in the 
 Startup folder, the system time is synchronized each time
 the computer is started. 

 Shortcut Path: "C:\Program Files\Time Sync\Time Sync.exe" /now

 You can add your local Network Time Server by modifying the
 SERVER.txt file, which can be found in the directorie where
 the program is installed.

 In some countrys there might be a problem with the date format
 and it's possible you must change to your local time format.
 This one works fine in Europe, but elsewere you could have
 31 months or 12 years, so check the date formats if errors

 *** Special thanks ***

 Special thanks to George McCoy and Jim Huff for the winsock code.
 (sorry for the late credits but found this code a long time ago
 and change the interface, added some stuff and used it for quit
 some time before submission. Unfortunally I lost track of the
 original submission. :-p