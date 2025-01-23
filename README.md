This project is me trying to recreate ADUC using PowerShell and Windows Presentation Framework and Windows Forms - since there is no ADUC through RSAT on ARM(atm).

Keeping the old versions here as well(v.0.1 is broken).

v0.7 is the newest iteration, GUI should scale according to screen size. No animation atm.
With this tool you will be able to:

- Create a single user and place it in an existing security group and OU, firstname.lastname@domain.com will be the username. Not customizeable at the moment.
- Create a Security Group, set the scope of it, and choose which OU to place it in.
- Create a Computer Object and which OU it should reside in.
- Create an Organizational Unit, and decide which existing OU it throw it into, including "root".
- Show the OU structure in a window, using XAML and a treeview. Shows the existing structure hierarchically.

Built-in Security Groups are filtered out, as well as built in CN's/OU's except "DC=domain,DC=local" when you create an Organizational Unit.

When you run the script, it will create a folder called "Data" in the project folder, which will contain .csv-files with exported information about existing OU's and Security Groups, as well as a .csv which holds the domain name written in under "Create a User".
It will also generate a .csv-file which retains the domain name entered in through the GUI. Writing something else overwrites the value in the file.
It'll also generate another xaml-file which holds the OU structure hierarchically, which is displayable in a new window.

NB! Running and exiting v.0.2 and v.0.3 script will kill the powershell process on the machine when closing out of the main script.
Not the case with v.0.4 and v.0.6.
