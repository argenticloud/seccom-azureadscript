<#
.SYNOPSIS
    Sync selected user attributes from on-prem AD to Entra ID (Microsoft Graph),
    report Exchange Online Archive status, and write Archive GUID back to on-prem AD.

.DESCRIPTION
    - Reads attributes from local AD (Get-ADUser)
    - Reads user from Entra ID (Get-MgUser)
    - Compares attributes and proposes AD updates interactively
    - Connects to Exchange Online to fetch Archive mailbox info (status, GUID, name)
    - Writes back EXO ArchiveGuid -> AD msExchArchiveGuid (binary) and ArchiveName -> msExchArchiveName
      (writeback occurs when -DryRun:$false)

.PARAMETER Target
    A UPN (e.g. user@domain.com) to process a single user, or "allusers" to process all AD users.

.PARAMETER DryRun
    When set, no updates are sent to Entra ID nor AD (report-only). (Default: On)

.PARAMETER WritebackArchive
    When set (default), evaluate archive writeback logic. With -DryRun:$false it will update AD.

.NOTES
    Requires:
      - ActiveDirectory module (for local AD)
      - Microsoft.Graph.Users module (Install-Module Microsoft.Graph -Scope CurrentUser)
      - ExchangeOnlineManagement module (Install-Module ExchangeOnlineManagement -Scope CurrentUser)
    Connects to Graph with scopes: User.ReadWrite.All, Directory.Read.All
    Connects to Exchange Online via Connect-ExchangeOnline

    AD REQUIREMENTS:
      - Exchange schema in on-prem AD (attributes like msExchArchiveGuid / msExchArchiveName)
      - Privileges to modify these attributes on user objects
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Target,

    [switch]$DryRun = $true,
    [switch]$WritebackArchive = $true,
    [switch]$PromptForChanges = $true
)

# --- Modules & Auth ---
#Requires -Modules ActiveDirectory
#Requires -Modules Microsoft.Graph.Users
#Requires -Modules ExchangeOnlineManagement

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Microsoft Graph (if not already connected) - read-only scope
try {
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All' | Out-Null
    }
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    return
}

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
} catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    return
}

# --- Report setup ---
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath   = ".\SyncReport_$timestamp.csv"
$report    = New-Object System.Collections.Generic.List[object]

# --- Attribute Map (AD -> Graph) ---
$AttributeMap = @{
    # Core identity/display
    'displayName'                  = @{ Graph='displayName';                  Writable=$true }
    'givenName'                    = @{ Graph='givenName';                    Writable=$true }
    'sn'                           = @{ Graph='surname';                      Writable=$true }              # AD sn -> Graph surname
    'userPrincipalName'            = @{ Graph='userPrincipalName';            Writable=$true }              # Be cautious changing UPN

    # Org info
    'department'                   = @{ Graph='department';                   Writable=$true }
    'title'                        = @{ Graph='jobTitle';                     Writable=$true }              # AD title -> Graph jobTitle
    'physicalDeliveryOfficeName'   = @{ Graph='officeLocation';               Writable=$true }

    # Contact
    'mobile'                       = @{ Graph='mobilePhone';                  Writable=$true; ReverseTransform = { param($gv) if ([string]::IsNullOrWhiteSpace($gv)) { $null } else { "$gv" } } }
    'telephoneNumber'              = @{ Graph='businessPhones';               Writable=$true; Transform = { param($v) if ([string]::IsNullOrWhiteSpace($v)) { @() } else { @("$v") } }; ReverseTransform = { param($gv) if ($gv -is [System.Collections.IEnumerable] -and -not ($gv -is [string])) { ($gv | Where-Object { $_ -and $_.Trim() -ne '' } | Select-Object -First 1) } elseif ([string]::IsNullOrWhiteSpace($gv)) { $null } else { "$gv" } } }
    'streetAddress'                = @{ Graph='streetAddress';                Writable=$true }
    'l'                            = @{ Graph='city';                         Writable=$true }              # AD l (locality) -> Graph city
    'st'                           = @{ Graph='state';                        Writable=$true }              # AD st -> Graph state
    'postalCode'                   = @{ Graph='postalCode';                   Writable=$true }
    'country'                      = @{ Graph='country';                      Writable=$true }

    # Read/Report-only (service or on-prem mastered)
    'mail'                         = @{ Graph='mail';                         Writable=$false }
    'proxyAddresses'               = @{ Graph='proxyAddresses';               Writable=$false }
    # On-prem extension attributes 1..15 – report only via Graph
    'extensionAttribute1'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute1'; Writable=$false }
    'extensionAttribute2'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute2'; Writable=$false }
    'extensionAttribute3'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute3'; Writable=$false }
    'extensionAttribute4'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute4'; Writable=$false }
    'extensionAttribute5'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute5'; Writable=$false }
    'extensionAttribute6'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute6'; Writable=$false }
    'extensionAttribute7'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute7'; Writable=$false }
    'extensionAttribute8'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute8'; Writable=$false }
    'extensionAttribute9'          = @{ Graph='onPremisesExtensionAttributes.extensionAttribute9'; Writable=$false }
    'extensionAttribute10'         = @{ Graph='onPremisesExtensionAttributes.extensionAttribute10'; Writable=$false }
    'extensionAttribute11'         = @{ Graph='onPremisesExtensionAttributes.extensionAttribute11'; Writable=$false }
    'extensionAttribute12'         = @{ Graph='onPremisesExtensionAttributes.extensionAttribute12'; Writable=$false }
    'extensionAttribute13'         = @{ Graph='onPremisesExtensionAttributes.extensionAttribute13'; Writable=$false }
    'extensionAttribute14'         = @{ Graph='onPremisesExtensionAttributes.extensionAttribute14'; Writable=$false }
    'extensionAttribute15'         = @{ Graph='onPremisesExtensionAttributes.extensionAttribute15'; Writable=$false }
}

# Helper: AD properties to retrieve (include archive attributes)
$AdPropsToGet = (@('UserPrincipalName','DistinguishedName','msExchArchiveGuid','msExchArchiveName') + $AttributeMap.Keys | Sort-Object -Unique)

# Graph properties to select (read-only)
$GraphPropsToSelect = @(
    'Id','UserPrincipalName','DisplayName','GivenName','Surname',
    'Department','JobTitle','Mail','MobilePhone','BusinessPhones','OfficeLocation',
    'City','State','PostalCode','Country','StreetAddress',
    'OnPremisesExtensionAttributes'
)

function Get-GraphPropertyValue {
    param([Parameter(Mandatory=$true)][object]$GraphUser,[Parameter(Mandatory=$true)][string]$GraphPath)
    $current = $GraphUser
    foreach ($segment in $GraphPath -split '\.') {
        if ($null -eq $current) { return $null }
        $current = $current | Select-Object -ExpandProperty $segment -ErrorAction SilentlyContinue
    }
    return $current
}

function Normalize-Value { param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value.Trim() }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $arr = @()
        foreach ($v in $Value) {
            if ($null -ne $v) { $s = "$v".Trim(); if ($s -ne '') { $arr += $s } }
        }
        return $arr
    }
    return $Value
}

function Equal-Values { param($A,$B)
    $aNorm = Normalize-Value -Value $A
    $bNorm = Normalize-Value -Value $B
    if ($null -eq $aNorm -and $null -eq $bNorm) { return $true }
    if ($null -eq $aNorm -or  $null -eq $bNorm) { return $false }
    if ($aNorm -is [System.Collections.IEnumerable] -and -not ($aNorm -is [string]) -and
        $bNorm -is [System.Collections.IEnumerable] -and -not ($bNorm -is [string])) {
        $aSet = @($aNorm | ForEach-Object { "$_".ToLowerInvariant() }) | Sort-Object
        $bSet = @($bNorm | ForEach-Object { "$_".ToLowerInvariant() }) | Sort-Object
        return ($aSet -join '|') -eq ($bSet -join '|')
    }
    if ($aNorm -is [string] -and $bNorm -is [string]) {
        return $aNorm.Equals($bNorm, [System.StringComparison]::OrdinalIgnoreCase)
    }
    return $aNorm -eq $bNorm
}

function Build-PatchObject { param([Parameter(Mandatory=$true)]$AdUser,[Parameter(Mandatory=$true)]$GraphUser)
    $patch = @{}
    $changes = @()
    foreach ($adName in $AttributeMap.Keys) {
        $map = $AttributeMap[$adName]
        $graphPath = $map.Graph
        $writable  = [bool]$map.Writable
        if (-not $writable) {
            $adVal = Normalize-Value ($AdUser.$adName)
            $cloudVal = Normalize-Value (Get-GraphPropertyValue -GraphUser $GraphUser -GraphPath $graphPath)
            if (-not (Equal-Values $adVal $cloudVal)) {
                $changes += [PSCustomObject]@{ Property=$graphPath; Before=($cloudVal|ConvertTo-Json -Compress -Depth 6); After=($adVal|ConvertTo-Json -Compress -Depth 6); Writable=$false }
            }
            continue
        }
        $adRaw = $AdUser.$adName
        $adDesired = if ($map.ContainsKey('Transform') -and $null -ne $map.Transform) { & $map.Transform $adRaw } else { $adRaw }
        $adDesired = Normalize-Value $adDesired
        $cloudCurrent = Normalize-Value (Get-GraphPropertyValue -GraphUser $GraphUser -GraphPath $graphPath)
        if (-not (Equal-Values $adDesired $cloudCurrent)) {
            $changes += [PSCustomObject]@{ Property=$graphPath; Before=($cloudCurrent|ConvertTo-Json -Compress -Depth 6); After=($adDesired|ConvertTo-Json -Compress -Depth 6); Writable=$true }
            if ($graphPath -notmatch '\.') { $patch[$graphPath] = $adDesired }
        }
    }
    [PSCustomObject]@{ PatchObject=$patch; Changes=$changes }
}

# Build a plan of AD updates using Graph as the source-of-truth for mapped attributes
function Build-AdUpdatePlan {
    param(
        [Parameter(Mandatory=$true)][Microsoft.ActiveDirectory.Management.ADUser]$AdUser,
        [Parameter(Mandatory=$true)]$GraphUser
    )
    $plan = @()
    foreach ($adName in $AttributeMap.Keys) {
        $map = $AttributeMap[$adName]
        $graphPath = $map.Graph
        $graphValRaw = Get-GraphPropertyValue -GraphUser $GraphUser -GraphPath $graphPath
        $graphValNorm = Normalize-Value $graphValRaw
        $desiredAd = $graphValNorm
        if ($map.ContainsKey('ReverseTransform') -and $null -ne $map.ReverseTransform) {
            $desiredAd = & $map.ReverseTransform $graphValNorm
        }
        $adCurrent = $AdUser.$adName
        if (-not (Equal-Values $adCurrent $desiredAd)) {
            $plan += [PSCustomObject]@{
                AdAttribute   = $adName
                GraphProperty = $graphPath
                AdCurrent     = $adCurrent
                EntraValue    = $graphValNorm
                AdDesired     = $desiredAd
            }
        }
    }
    return ,$plan
}

function Apply-AdUpdatesWithPrompt {
    param(
        [Parameter(Mandatory=$true)][Microsoft.ActiveDirectory.Management.ADUser]$AdUser,
        [Parameter(Mandatory=$true)][object[]]$Plan,
        [switch]$DoWrite,
        [switch]$Prompt
    )
    $applied = @()
    foreach ($chg in $Plan) {
        $entraDispl = ($chg.EntraValue | ConvertTo-Json -Compress -Depth 6)
        $adDispl    = ($chg.AdCurrent  | ConvertTo-Json -Compress -Depth 6)
        $msg = "User $($AdUser.UserPrincipalName) has entra $($chg.GraphProperty) of $entraDispl and ad $($chg.AdAttribute) of $adDispl. Update local AD? y/n"
        $doIt = $false
        if ($Prompt) {
            $answer = Read-Host $msg
            if ($answer -match '^[Yy]') { $doIt = $true }
        } else {
            $doIt = $true
        }
        if ($doIt) {
            if ($DoWrite) {
                try {
                    if ($null -eq $chg.AdDesired -or ("$($chg.AdDesired)".Trim() -eq '')) {
                        Set-ADUser -Identity $AdUser.DistinguishedName -Clear $chg.AdAttribute -ErrorAction Stop
                    } else {
                        try {
                            Set-ADUser -Identity $AdUser.DistinguishedName -Replace @{ ($chg.AdAttribute) = $chg.AdDesired } -ErrorAction Stop
                        } catch {
                            Set-ADUser -Identity $AdUser.DistinguishedName -Add @{ ($chg.AdAttribute) = $chg.AdDesired } -ErrorAction Stop
                        }
                    }
                    $applied += $chg
                } catch {
                    Write-Warning "Failed to update AD attribute $($chg.AdAttribute) for $($AdUser.UserPrincipalName): $($_.Exception.Message)"
                }
            } else {
                # Dry run: treat as applied for reporting purposes
                $applied += $chg
            }
        }
    }
    return ,$applied
}

# --- Archive GUID helpers ---
function Convert-GuidStringToByteArray {
    param([Parameter(Mandatory=$true)][string]$GuidString)
    try {
        $g = [Guid]::Parse($GuidString)
        return $g.ToByteArray()  # little-endian as required by AD for *msExch* GUID attrs
    } catch {
        return $null
    }
}

function Convert-ByteArrayToGuidString {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -ne 16) { return $null }
    try {
        return (New-Object Guid (,$Bytes)).ToString()
    } catch { return $null }
}

function Compare-ByteArrays {
    param([byte[]]$A,[byte[]]$B)
    if ($null -eq $A -and $null -eq $B) { return $true }
    if ($null -eq $A -or  $null -eq $B) { return $false }
    if ($A.Length -ne $B.Length) { return $false }
    for ($i=0; $i -lt $A.Length; $i++) { if ($A[$i] -ne $B[$i]) { return $false } }
    return $true
}

function Writeback-ArchiveToAD {
    param(
        [Parameter(Mandatory=$true)][Microsoft.ActiveDirectory.Management.ADUser]$AdUser,
        [Parameter(Mandatory=$true)]$ExoMailbox,
        [switch]$DoWrite,
        [switch]$Prompt
    )
    $wb = [ordered]@{
        Attempted             = [bool]$DoWrite
        ArchiveStatus         = $null
        ExoArchiveGuidString  = $null
        AdArchiveGuidBefore   = $null
        AdArchiveGuidAfter    = $null
        AdArchiveNameBefore   = $null
        AdArchiveNameAfter    = $null
        Action                = 'None'
        Error                 = ''
    }

    # Gather EXO archive
    if (-not $ExoMailbox) { $wb.Error = 'EXO mailbox not found'; return [PSCustomObject]$wb }
    $wb.ArchiveStatus        = "$($ExoMailbox.ArchiveStatus)"   # 'Active' or 'None'
    $wb.ExoArchiveGuidString = if ($ExoMailbox.ArchiveGuid) { "$($ExoMailbox.ArchiveGuid)" } else { '' }

    # Current AD values
    $adGuidBytes   = $AdUser.'msExchArchiveGuid'
    $adNameCurrent = $AdUser.'msExchArchiveName'
    $wb.AdArchiveGuidBefore = Convert-ByteArrayToGuidString -Bytes $adGuidBytes
    $wb.AdArchiveNameBefore = if ($null -ne $adNameCurrent) { "$adNameCurrent" } else { $null }

    # Decide desired
    $desiredGuidBytes = $null
    $desiredName      = $null
    if ($ExoMailbox.ArchiveStatus -eq 'Active' -and [string]::IsNullOrEmpty($wb.ExoArchiveGuidString) -eq $false) {
        $desiredGuidBytes = Convert-GuidStringToByteArray -GuidString $wb.ExoArchiveGuidString
        $desiredName      = if ($ExoMailbox.ArchiveName) { "$($ExoMailbox.ArchiveName)" } else { $null }
    } else {
        # Archive not active: no destructive clear by default; keep AD as-is
        $wb.Action = if ($wb.ArchiveStatus -eq 'Active') { 'NoChange' } else { 'ArchiveNotActive-ReportedOnly' }
        $wb.AdArchiveGuidAfter = $wb.AdArchiveGuidBefore
        $wb.AdArchiveNameAfter = $wb.AdArchiveNameBefore
        return [PSCustomObject]$wb
    }

    # Compare and (optionally) write
    $guidNeedsUpdate = -not (Compare-ByteArrays -A $adGuidBytes -B $desiredGuidBytes)
    $nameNeedsUpdate = $false
    if ($desiredName) {
        $nameNeedsUpdate = -not (Equal-Values $adNameCurrent $desiredName)
    }

    if (-not $guidNeedsUpdate -and -not $nameNeedsUpdate) {
        $wb.Action = 'NoChange'
        $wb.AdArchiveGuidAfter = $wb.AdArchiveGuidBefore
        $wb.AdArchiveNameAfter = $wb.AdArchiveNameBefore
        return [PSCustomObject]$wb
    }

    if (-not $DoWrite) {
        $wb.Action = 'WouldUpdate'
        $wb.AdArchiveGuidAfter = Convert-ByteArrayToGuidString -Bytes $desiredGuidBytes
        $wb.AdArchiveNameAfter = $desiredName
        return [PSCustomObject]$wb
    }

    # Perform updates
    try {
        if ($Prompt) {
            $displayGuidBefore = $wb.AdArchiveGuidBefore
            $displayGuidAfter  = Convert-ByteArrayToGuidString -Bytes $desiredGuidBytes
            $promptMsg = "User $($AdUser.UserPrincipalName) has entra OnlineArchive of $($wb.ArchiveStatus) (Guid $displayGuidAfter) and ad msExchArchiveGuid of $displayGuidBefore. Update local AD? y/n"
            $ans = Read-Host $promptMsg
            if ($ans -notmatch '^[Yy]') { $wb.Action = 'SkippedByUser'; return [PSCustomObject]$wb }
        }
        $replace = @{}
        if ($guidNeedsUpdate) { $replace['msExchArchiveGuid'] = $desiredGuidBytes }
        if ($nameNeedsUpdate) { $replace['msExchArchiveName'] = $desiredName }

        if ($replace.Count -gt 0) {
            try {
                Set-ADUser -Identity $AdUser.DistinguishedName -Replace $replace -ErrorAction Stop
            } catch {
                # If -Replace fails (attribute missing), try -Add
                Set-ADUser -Identity $AdUser.DistinguishedName -Add $replace -ErrorAction Stop
            }
        }

        # Re-read current to confirm
        $fresh = Get-ADUser -Identity $AdUser.DistinguishedName -Properties msExchArchiveGuid,msExchArchiveName
        $wb.AdArchiveGuidAfter = Convert-ByteArrayToGuidString -Bytes $fresh.'msExchArchiveGuid'
        $wb.AdArchiveNameAfter = if ($fresh.'msExchArchiveName') { "$($fresh.'msExchArchiveName')" } else { $null }
        $wb.Action = 'Updated'
    } catch {
        $wb.Action = 'Failed'
        $wb.Error  = "Archive writeback failed: $($_.Exception.Message)"
        # still set After to intended view
        $wb.AdArchiveGuidAfter = Convert-ByteArrayToGuidString -Bytes $desiredGuidBytes
        $wb.AdArchiveNameAfter = $desiredName
    }

    return [PSCustomObject]$wb
}

function Sync-User {
    param([Parameter(Mandatory=$true)][string]$Upn)

    $errors = New-Object System.Collections.Generic.List[string]

    # --- Get local AD user ---
    $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$Upn'" -Properties $AdPropsToGet -ErrorAction SilentlyContinue
    if (-not $adUser) {
        Write-Warning "AD user not found: $Upn"
        return [PSCustomObject]@{
            UserPrincipalName     = $Upn
            DirsyncUser           = $null
            DryRun                = [bool]$DryRun
            ForceCloudWrite       = [bool]$ForceCloudWrite
            ActionTaken           = 'NotFound-AD'
            ChangeCount           = 0
            WritableChangeCount   = 0
            ChangesJson           = '[]'
            PatchJson             = '{}'
            ArchiveMailboxEnabled = $null
            ArchiveStatus         = 'Unknown'
            ArchiveGuid           = ''
            ArchiveName           = ''
            ArchiveWritebackJson  = '{}'
            Error                 = 'AD user not found'
        }
    }

    # --- Get Graph user ---
    $graphUser = $null
    try { $graphUser = Get-MgUser -UserId $Upn -Property $GraphPropsToSelect -ErrorAction Stop }
    catch {
        $msg = "Graph user not found or inaccessible: $Upn - $($_.Exception.Message)"
        Write-Warning $msg
        return [PSCustomObject]@{
            UserPrincipalName     = $Upn
            DirsyncUser           = $null
            DryRun                = [bool]$DryRun
            ForceCloudWrite       = [bool]$ForceCloudWrite
            ActionTaken           = 'NotFound-Graph'
            ChangeCount           = 0
            WritableChangeCount   = 0
            ChangesJson           = '[]'
            PatchJson             = '{}'
            ArchiveMailboxEnabled = $null
            ArchiveStatus         = 'Unknown'
            ArchiveGuid           = ''
            ArchiveName           = ''
            ArchiveWritebackJson  = '{}'
            Error                 = $msg
        }
    }

    $dirsync = $null

    # --- Build change summary (read-only) ---
    $result   = Build-PatchObject -AdUser $adUser -GraphUser $graphUser
    $patch    = $result.PatchObject
    $changes  = $result.Changes
    $wChanges = $changes | Where-Object { $_.Writable -eq $true }

    # --- Graph -> AD interactive writeback for mapped attributes ---
    $adPlan = Build-AdUpdatePlan -AdUser $adUser -GraphUser $graphUser
    $adApplied = Apply-AdUpdatesWithPrompt -AdUser $adUser -Plan $adPlan -DoWrite:(!$DryRun) -Prompt:$PromptForChanges

    # --- Exchange Online: Archive status ---
    $archiveEnabled = $null
    $archiveStatus  = 'Unknown'
    $archiveGuidStr = ''
    $archiveNameStr = ''
    $exoMailbox = $null
    try {
        $exoMailbox = Get-Mailbox -Identity $Upn -ErrorAction Stop -WarningAction SilentlyContinue
        if ($exoMailbox) {
            $archiveStatus  = "$($exoMailbox.ArchiveStatus)"      # 'None' or 'Active'
            $archiveEnabled = $exoMailbox.ArchiveStatus -eq 'Active'
            $archiveGuidStr = if ($exoMailbox.ArchiveGuid) { "$($exoMailbox.ArchiveGuid)" } else { '' }
            $archiveNameStr = if ($exoMailbox.ArchiveName) { "$($exoMailbox.ArchiveName)" } else { '' }
        }
    } catch {
        $errors.Add("EXO Get-Mailbox failed: $($_.Exception.Message)") | Out-Null
    }

    # --- Archive writeback to AD ---
    $wbDetails = [PSCustomObject]@{ }
    if ($WritebackArchive) {
        $wbDetails = Writeback-ArchiveToAD -AdUser $adUser -ExoMailbox $exoMailbox -DoWrite:(!$DryRun) -Prompt:$PromptForChanges
    } else {
        $wbDetails = [PSCustomObject]@{ Attempted=$false; Action='Skipped'; Error=''; ArchiveStatus=$archiveStatus; ExoArchiveGuidString=$archiveGuidStr; AdArchiveGuidBefore=(Convert-ByteArrayToGuidString -Bytes $adUser.'msExchArchiveGuid'); AdArchiveGuidAfter=$null; AdArchiveNameBefore=$adUser.'msExchArchiveName'; AdArchiveNameAfter=$null }
    }

    # --- Build report row ---
    return [PSCustomObject]@{
        UserPrincipalName       = $Upn
        DirsyncUser             = $dirsync
        DryRun                  = [bool]$DryRun
        ActionTaken             = if ($wChanges.Count -eq 0) { 'NoChange' } else { 'ReportedOnly' }
        ChangeCount             = $changes.Count
        WritableChangeCount     = $wChanges.Count
        ChangesJson             = ($changes | ConvertTo-Json -Compress -Depth 6)
        PatchJson               = ($patch   | ConvertTo-Json -Compress -Depth 6)
        AdPlannedChanges        = ($adPlan | ConvertTo-Json -Compress -Depth 6)
        AdAppliedChanges        = ($adApplied | ConvertTo-Json -Compress -Depth 6)
        ArchiveMailboxEnabled   = $archiveEnabled
        ArchiveStatus           = $archiveStatus
        ArchiveGuid             = $archiveGuidStr
        ArchiveName             = $archiveNameStr
        ArchiveWritebackJson    = ($wbDetails | ConvertTo-Json -Compress -Depth 6)
        Error                   = if ($errors.Count -gt 0) { ($errors -join ' | ') } else { '' }
    }
}

# --- Main ---
if ($Target -ieq 'allusers') {
    Write-Host "Processing all AD users..." -ForegroundColor Cyan
    $all = Get-ADUser -Filter * -Properties $AdPropsToGet
    $i = 0
    foreach ($u in $all) {
        $i++
        Write-Progress -Activity "Sync users" -Status "$i / $($all.Count): $($u.UserPrincipalName)" -PercentComplete (($i / [math]::Max(1,$all.Count)) * 100)
        $row = Sync-User -Upn $u.UserPrincipalName
        $report.Add($row) | Out-Null
    }
    Write-Progress -Activity "Sync users" -Completed
} else {
    $row = Sync-User -Upn $Target
    $report.Add($row) | Out-Null
}

# --- Export report ---
$report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Sync complete. Report saved to $csvPath" -ForegroundColor Green
