# syncfix.ps1


Pulls user info from AD, Entra (Graph), and Exchange Online
Compares AD → Entra attributes (based on a mapping in the script)
Prompts per attribute and writes to local AD only (never writes to Entra)
Reports archive mailbox status (Active/None, GUID, Name)
Writes Archive GUID and Archive Name back to AD (prompted)
Exports a CSV report with all the details

Basically what I changed was removing trying to write to Entra (as we are using local 
AD only, as the source of truth but mapping entra-created values) was add a flag so 
that you can confirm Y/N when it runs whether to update the value.

Prerequisites:
Windows PowerShell 5.1

Modules:

  Install-Module ActiveDirectory
  
  Install-Module Microsoft.Graph
  
  Install-Module ExchangeOnlineManagement


Permissions:

  Graph: User.Read.All, Directory.Read.All (read-only)
  
  EXO: Get-Mailbox rights
  
  AD: Can update msExchArchiveGuid and msExchArchiveName (and any mapped user attributes)
  

  ** This script never writes to Entra. It reads Graph and EXO only to compare values. **


  Instructions to run:
  
  **REPORT ONLY <--- USE THIS ONE FIRST**
  
  .\syncfix.ps1 -Target user@domain.com

  **All Users (REPORT ONLY) <--- USE THIS ONE SECOND**
  
  .\syncfix.ps1 -Target allusers

  **Update local AD with prompts - SINGLE USER <--- USE THIS ONE THIRD**
  
  .\syncfix.ps1 -Target user@domain.com -DryRun:$false

  **All Users (UPDATE AD with prompts) <-- USE THIS ONE LAST - it will go through everything**
  
  .\syncfix.ps1 -Target allusers -DryRun:$false


 **Skip Archive Writeback (kindof pointless since that's our issue, but for completeness...)**
  
  .\syncfix.ps1 -Target user@domain.com -WritebackArchive:$false


  **Disable per-attribute prompts (apply all mapped changes silently - don't use this one, it's**
  **there for completeness and in case we want to use this later for someone else later and we**
  **are confident it won't break things)**
  
  .\syncfix.ps1 -Target user@domain.com -DryRun:$false -PromptForChanges:$false


  You should get a CSV file to review what has been changed.

  
  
  Contact me if you have issues.
