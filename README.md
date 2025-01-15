This project is me trying to recreate ADUC using PowerShell and Windows Presentation Framework and Windows Forms - since there is no ADUC through RSAT on ARM(atm).

ADUC-PS.ps1 is the script to launch, which will in turn launch the module selected using the GUI. When running this script:

  - (only on v.0.4) Creates a directory for the .csv-files named "Data" in the project folder.
  - Exports OU names and distinguished names into a .csv-file.
  - Exports Security groups and stores it in a .csv-file(no filter, yet)
  - Creates a .csv-file which will save what is written into the "domain"-textbox and populate it. Will overwrite if something else is written in.

It will not create a .csv for creating users in bulk(for now at least)

Files are stored under "Modules\Data" in the project folder(just "Data" in v.0.4). 
Users created will combine first and last name into a username, e.g. "firstname.lastname@domain.com". No flexibility there atm.

Workin versions are 0.2, 0.3 and 0.4. Keeping the older versions in here, just for the heck of it.
Version 0.4 is now one single script instead - a bit smoother to use IMO. Such animation, much fancy :D

NB! Running and exiting this script will kill the powershell process on the machine when closing out of the main script (not relevant on 0.4, since it's not comprised of modules).
