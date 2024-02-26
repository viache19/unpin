# unpin
Unpin apps from the Windows Task Bar

This started because of the new company's policy, after each reboot the office aplications as Excel, Word and Outlook were added to my taskbar by default, and I do not really like when somebody touches my taskbar because with that changes all the already pinned applications where been moved.

So I created this little script. Basically what it will do, is to go to the shortcut located in "C:\ProgramData\Microsoft\Windows\Start Menu\Programs", make a right click and select the "Unpin from taskbar" option

Dependencies : 
  - pywin32 : pip install pywin32

Add to the autoruns registry : 

  - Press the key "Win + R", write the regedit
  - Right click, New, String Value
  - Name : the one you want
  - Value : python "C:\the\path\to\your\script.py"
