This project is me trying to recreate ADUC using PowerShell and Windows Forms, since there is no ADUC through RSAT on ARM(atm).

Now a v0.3 with some XML-styling on ADUC-PS.ps1 as well as SecGroup-Module.ps1 & CreateComputer-Module.ps1.

NB! Running and exiting this script will kill the powershell process on the machine when closing out of the main script.

ADUC-PS.ps1 is the script to launch, which will in turn launch the module selected using the GUI. When running this script:

  - Exports OU names and distinguished names into a .csv-file.
  - Exports Security groups and stores it in a .csv-file(no filter, yet)
  - Creates a .csv-file which will save what is written into the "domain"-textbox and populate it. Will overwrite if something else is written in.

It will not create a .csv for creating users in bulk(for now at least)

Files are stored under "Modules\Data" in the project folder. Users created will combine first and last name into a username, e.g. "kevin.silden@domain.com". No flexibility there atm.

Tested on a VM running Windows Server 2022 Standard Ed.
