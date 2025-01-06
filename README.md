This project is me trying to recreate ADUC using PowerShell and Windows Forms, since there is no ADUC through RSAT on ARM(atm).

NB! Running and exiting this script will kill the powershell process on the machine when closing out of the main script.

ADUC-PS.ps1 is the script to launch, which will in turn launch the module selected using the GUI.
When running this script, it will export the OU's available into a .csv-file under "Modules\Data\". 
The other modules will then use the data in this file for selecting OU's to store the objects created in AD.

For now, only the CreateUser-Module.ps1-script is working as intended. Tested on a VM running Windows Server 2022 Standard Ed.
The rest is just older scripts thrown in there and I haven't tested it yet.

- Kevin Silden
