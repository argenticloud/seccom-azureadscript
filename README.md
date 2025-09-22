# syncfix.ps1


Pulls user info from AD, Entra (Graph), and Exchange Online
Compares AD → Entra attributes (based on a mapping in the script)
Updates Entra (optional)
Reports archive mailbox status (Active/None, GUID, Name)
Writes Archive GUID and Archive Name back to AD (optional)
Exports a CSV report with all the details

Prerequisites:
Windows PowerShell 5.1

Modules:
  Install-Module ActiveDirectory
  Install-Module Microsoft.Graph
  Install-Module ExchangeOnlineManagement

Permissions:
  Graph: User.ReadWrite.All, Directory.Read.All
  EXO: Get-Mailbox rights
  AD: Can update msExchArchiveGuid and msExchArchiveName

  ** I hate asking for global admin here, but it might be relevant if there's no way to scope a lesser user. **

  Instructions to run:
  **REPORT ONLY**
  .\syncfix.ps1 -Target user@domain.com

  **Update Entra + write archive back to AD - SINGLE USER**
  .\syncfix.ps1 -Target user@domain.com -DryRun:$false

  **All Users (REPORT ONLY)**
  .\syncfix.ps1 -Target allusers

  **Skip Archive Writeback (kindof pointless since that's our issue, but for completeness...)**
  .\syncfix.ps1 -Target user@domain.com -WritebackArchive:$false

  --add the switch **-ForceCloudWrite** to the end of the command in order to work with Entra ID Connect running. 
  so ** .\syncfix.ps1 -Target allusers -DryRun:$false -ForceCloudWrite** is the final command to do it all.

  You should get a CSV file to review what has been changed.
  Contact me if you have issues.
