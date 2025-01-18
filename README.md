This project is me trying to recreate ADUC using PowerShell and Windows Presentation Framework and Windows Forms - since there is no ADUC through RSAT on ARM(atm).

v0.6 is the newest iteration, GUI should scale according to screen size. I noticed that the titles wouldn't nessecary scale when the window is sized down to a certain scale.
Workin versions are 0.2, 0.3, 0.4 and 0.6. Keeping the older versions in here, just for the heck of it.

When you run the script, it will create a folder called "Data" in the project folder, which will contain .csv-files which gathers information about OU's and Security Groups.
It will also generate a .csv-file which retains the domain name entered in through the GUI. Writing something else overwrites the value in the file.

NB! Running and exiting v.0.2 and v.0.3 script will kill the powershell process on the machine when closing out of the main script.
Not the case with v.0.4 and v.0.6.
