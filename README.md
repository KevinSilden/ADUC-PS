This project is me trying to recreate ADUC using PowerShell and Windows Forms, since there is no ADUC through RSAT on ARM(atm).

NB! Running and exiting this script will kill the powershell process on the machine when closing out of the main script.

ADUC-PS.ps1 is the script to launch, which will in turn launch the module selected using the GUI.
When running this script:
  - Exports OU names and distinguished names into a .csv-file.
  - Exports Security groups and stores it in a .csv-file.
  - Creates a .csv-file which will save what is written into the "domain"-textbox and populate it. Will overwrite if something else is written in.
Files are stored under "Modules\Data". Files added to this repo are from a VM for testing.
Users created will combine first and last name into a username, e.g. "kevin.silden@domain.com". No flexibility there atm.

Modules that are now in a working state:
- CreateUser-Module.ps1
- SecGroup-Module.ps1-script
- CreateComputer-Module.ps1

  The CSV-Module.ps1 is somewhat basic atm, needs rewriting.

Tested on a VM running Windows Server 2022 Standard Ed.
