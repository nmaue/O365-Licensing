# O365-Licensing
License Office 365 users according to on-premises AD OUs  
  
This script can be customized and run as a scheduled task on a server connected to an Active Directory forest.  
  
The script gets Origanizational Units from the AD forest and assigns them the correct license. 
It only assigns licenses to enabled users.  
  
Licenses can also be removed from an entire OU if defined

## Todo
* Change the piping to actual variable statements
* Clean it up using powershell classes
