--Applescript to fix scroll-lock problem with Microsoft Excel on Mac OS
--Written by Damien Clark (http://damosworld.wordpress.com)
--Licenced under GPLv2 (http://www.gnu.org/licenses/gpl-2.0.html)
--18th of June, 2012
--Updated: 22nd of July, 2015

--From time to time you may find that using the cursor keys in MS Excel for Mac will scroll the worksheet, but not shift the selected cell.  This is a feature of Excel that is enabled by pressing the Scroll Lock key.  On Macintosh computers once upon a time, F14 was the equivalent key to Scroll lock on the PC.  Modern Macintosh computers are no longer furnished with an F14 key, and while I am sure there is a keyboard combination to emulate F14, I have yet to find it.  As a result, you can find yourself stuck with this special scroll lock mode of Excel and it is very frustrating.  This script effectively "sends" an F14 keypress to Microsoft Excel, thus turning on or off this scroll lock feature.  So if you find yourself in this situation, you can execute this program to toggle the feature on and off as necessary.  Simply run the application and click OK.  

set returnedItems to (display dialog "Press OK to send scroll lock keypress to Microsoft Excel or press Quit" with title "Excel Scroll-lock Fix" buttons {"Quit", "OK"} default button 2)
set buttonPressed to the button returned of returnedItems

if buttonPressed is "OK" then
	tell application "Microsoft Excel"
		activate --Make it the active application to receive System keyboard events
	end tell
	tell application "System Events"
		key code 107 using {shift down} --Send the shift F14 key which is scroll-lock on Mac keyboards
	end tell
	activate
	display dialog "Scroll Lock key sent to Microsoft Excel" with title "Mac Excel Scroll-lock Fix" buttons {"OK"}
end if
