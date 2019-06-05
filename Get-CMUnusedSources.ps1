<#
.SYNOPSIS

.DESCRIPTION

.INPUTS

.EXAMPLE

.NOTES
    FileName:
    Authors: 
    Contact:
    Created:
    Updated:
    Thanks to:
        Cody Mathis (Windows Admin slack)
        Chris Kibble (Windows Admin slack)
        Chris Dent (Windows Admin slack)
        PsychoData ("the regex mancer", Windows Admin slack)
        Patrick (Windows Admin slack)
    
    Version history:
    1

    TODO: 
        - Comment regex
        - Review comments
        - Dashimo, ultimatedashboard?
        - Logging
        - Any functions accessing variables in parent scope and not passed as parameter to said function? Clean it!
            - Get-AllFolders for -AltFolderSearch
        - Application DT uninstall content source locations
        - Consider supporting \\?\UNC\server\share format, allows you to avoid low path limit
        - Have I stupidly assumed share name is same as folder name on disk???

    Test plan:
        - content objects with:
            - no content path
            - Local paths
            - Permission denied on various folders with content object source paths inside
            - server unreachable
            - server reachable but share doesn't exist
            - server reachable, share exists but no longer a valid path
            - 2 (or more) shared folders pointing to same path
            - Running from site server and specifying local path for $SourcesLocation

#>
#Requires -Version 5.1
[cmdletbinding(DefaultParameterSetName='1')]
Param (
    [Parameter(
        ParameterSetName='1',
        Mandatory=$true, 
        Position = 0
    )]
    [Parameter(
        ParameterSetName='2',
        Mandatory=$true, 
        Position = 0
    )]
    [ValidateScript({
        If (!([System.IO.Directory]::Exists($_))) {
            Throw "Invalid path or access denied"
        } ElseIf (!($_ | Test-Path -PathType Container)) {
            Throw "Value must be a directory, not a file"
        } Else {
            return $true
        }
    })]
    [string]$SourcesLocation,
    [Parameter(
        ParameterSetName='1',
        Mandatory=$true, 
        Position = 1
    )]
    [Parameter(
        ParameterSetName='2',
        Mandatory=$true, 
        Position = 1
    )]
    [ValidatePattern('^[a-zA-Z0-9]{3}$')]
    [string]$SiteCode,
    [Parameter(
        ParameterSetName='1',
        Mandatory=$true, 
        Position = 2
    )]
    [Parameter(
        ParameterSetName='2',
        Mandatory=$true, 
        Position = 2
    )]
    [string]$SCCMServer,
    [Parameter(
        ParameterSetName='1'
    )]
    [switch]$All = $true,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$Packages,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$Applications,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$Drivers,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$DriverPackages,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$OSImages,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$OSUpgradeImages,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$BootImages,
    [Parameter(
        ParameterSetName='2'
    )]
    [switch]$DeploymentPackages,
    [switch]$AltFolderSearch,
    [switch]$NoProgress
)

Function Get-CMContent {
    Param(
        $Commands,
        [string]$SCCMServer
    )
    # Invaluable resource for getting all source locations: https://www.verboon.info/2013/07/configmgr-2012-script-to-retrieve-source-path-locations/
    $AllContent = @()
    $ShareCache = @{}
    ForEach ($Command in $Commands) {
        ForEach ($item in (Invoke-Expression $Command)) {
            switch ($Command) {
                "Get-CMApplication" {
                    $AppMgmt = ([xml]$item.SDMPackageXML).AppMgmtDigest
                    $AppName = $AppMgmt.Application.DisplayInfo.FirstChild.Title
                    ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
                        $SourcePath = $DeploymentType.Installer.Contents.Content.Location
                        $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SCCMServer $SCCMServer
                        $obj = New-Object PSObject
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value "$($DeploymentType.AuthoringScopeId)/$($DeploymentType.LogicalName)"
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value "$($item.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                        $AllContent += $obj
                        $ShareCache = $GetAllPathsResult[0]
                    }
                }
                "Get-CMDriver" { # I don't actually think it's possible for a driver to not have source path set
                    $SourcePath = $item.ContentSourcePath
                    $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SCCMServer $SCCMServer    
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.CI_ID
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.LocalizedDisplayName
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                    $AllContent += $obj
                    $ShareCache = $GetAllPathsResult[0]
                }
                default {
                    # OS images and boot iamges are absolute paths to files
                    If ("Get-CMOperatingSystemImage","Get-CMBootImage" -contains $Command) {
                        $SourcePath = Split-Path $item.PkgSourcePath
                    }
                    Else {
                        $SourcePath = $item.PkgSourcePath
                    }
                    $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SCCMServer $SCCMServer
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.PackageId
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.Name
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                    $AllContent += $obj
                    $ShareCache = $GetAllPathsResult[0]
                }   
            }
        }
    }
    return $AllContent
}

Function Get-AllPaths {
    param (
        [string]$Path,
        [hashtable]$Cache,
        [string]$SCCMServer
    )

    $AllPaths = @{}
    $result = @()

    If ([string]::IsNullOrEmpty($Path) -eq $false) {
        $Path = $Path.TrimEnd("\")
    }

    ##### Determine path type

    switch ($true) {
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)(\\[a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)") {
            # Path that is \\server\share\folder
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$Matches[3]
        }
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)$") {
            # Path that is \\server\share
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$null
        }
        ($Path -match "^[a-zA-Z]:\\") {
            # Path that is local
            $AllPaths.Add($Path, $SCCMServer)
            $result += $Cache
            $result += $AllPaths
            return $result
        }
        ([string]::IsNullOrEmpty($Path) -eq $true) {
            $result += $Cache
            $result += $AllPaths
            return $result
        }
        default { 
            Write-Warning "Unable to interpret path `"$($Path)`", used by `"$($obj.Name)`""
            $AllPaths.Add($Path, $null)
            $result += $Cache
            $result += $AllPaths
            return $result
        }
    }

    ##### Determine FQDN, IP and NetBIOS

    If (Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue) {
        If ($Server -as [IPAddress]) {
            try {
                $FQDN = [System.Net.Dns]::GetHostEntry($Server).HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                $FQDN = $null
            }
            $IP = $Server
        }
        Else {
            try {
                $FQDN = [System.Net.Dns]::GetHostByName($Server).HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                $FQDN = $null
            }
            $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
        }
    }
    Else {
        Write-Warning "Server `"$($Server)`" is unreachable, used by `"$($obj.Name)`""
        $AllPaths.Add($Path, $null)
        $result += $Cache
        $result += $AllPaths
        return $result
    }

    ##### Update the cache of shared folders and their local paths

    If (($Cache.Keys -contains $FQDN) -eq $false) {
        # Do not yet have this server's shares cached
        # $AllSharedFolders is null if couldn't connect to serverr to get all shared folders
        $AllSharedFolders = Get-AllSharedFolders -Server $FQDN
        If ([string]::IsNullOrEmpty($AllSharedFolders) -eq $false) {
            $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object {
                $Cache.Add($_, $AllSharedFolders)
            }
        }
        Else {
            Write-Warning "Could not update cache because could not get shared folders from: `"$($FQDN)`", used by `"$($obj.Name)`""
        }
    }

    ##### Build the AllPaths property

    $AllPathsArr = @()

    ## Build AllPaths based on share name from given Path

    $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object -Process {
        $AltServer = $_
        $LocalPath = ($Cache.$AltServer.GetEnumerator() | Where-Object { $_.Key -eq $ShareName }).Value
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            $AllPathsArr += ("\\$($AltServer)\$($LocalPath)$($ShareRemainder)" -replace ':', '$')
            $SharesWithSamePath = ($Cache.$AltServer.GetEnumerator() | Where-Object { $_.Value -eq $LocalPath }).Key
            $SharesWithSamePath | ForEach-Object -Process {
                $AltShareName = $_
                $AllPathsArr += "\\$($AltServer)\$($AltShareName)$($ShareRemainder)"
            }
        }
        Else {
            Write-Warning "Share `"$($ShareName)`" does not exist on `"$($_)`", used by `"$($obj.Name)`""
        }
        $AllPathsArr += "\\$($AltServer)\$($ShareName)$($ShareRemainder)"
    } -End {
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            If ($LocalPath -match "^[a-zA-Z]:$") {
                $AllPathsArr += "$($LocalPath)\"
            }
            Else {
                $AllPathsArr += "$($LocalPath)$($ShareRemainder)"
            }
        }
    }

    ForEach ($item in $AllPathsArr) {
        If (($AllPaths.Keys -contains $item) -eq $false) {
            $AllPaths.Add($item, $NetBIOS)
        }
    }

    $result += $Cache
    $result += $AllPaths
    return $result
}

Function Get-AllSharedFolders {
    Param([String]$Server)

    $AllShares = @{}

    try {
        $Shares = Get-WmiObject -ComputerName $Server -Class Win32_Share -ErrorAction Stop | Where-Object {-not [string]::IsNullOrEmpty($_.Path)}
        ForEach ($Share in $Shares) {
            # The TrimEnd method is only really concerned for drive letter shares
            # as they're usually stored as f$ = "F:\" and this messes up Get-AllPaths a little
            $AllShares += @{ $Share.Name = $Share.Path.TrimEnd("\") }
        }
    }
    catch {
        $AllShares = $null
    }

    return $AllShares
}

Function Get-AllFolders {
    Param(
        [string]$Path,
        [bool]$AltFolderSearch
    )


    switch ($true) { 
        ($Path -match "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z0-9\\`~!@#$%^&(){}\'._ -]+") {
            # Matches if it's a UNC path
            # Could have queried .IsUnc property on [System.Uri] object but I wanted to verify user hadn't first given us \\?\ path type
            $Path = $Path -replace "^\\\\\?\\UNC\\", "\\"
            break
        }
        ($Path -match "^[a-zA-Z]:\\") {
            $Path = "\\?\" + $Path
            break
        }
        default {
            Write-Warning "Couldn't determine path type for `"$Paths`" so might have problems accessing folders that breach MAX_PATH limit"
        }
    }
    
    If ($Path -match "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z0-9\\`~!@#$%^&(){}\'._ -]+") {
        $Path = $Path -replace "^\\\\\?\\UNC\\", "\\"
    }
    
    If ($AltFolderSearch) {
        [System.Collections.ArrayList]$result = Start-AltFolderSearch -FolderName $Path
    }
    Else {
        try {
            [System.Collections.ArrayList]$result = (Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue).FullName
        }
        catch {
            Throw "Consider using -AltFolderSearch"
        }
    }
    
    switch ($true) {
        ($Path -match "^\\\\\?\\UNC\\") {
            $Path = $Path -replace [regex]::Escape("\\?\UNC"), "\"
            $result.Add($Path)
            $result = $result -replace [Regex]::Escape("\\?\UNC"), "\"
            break
        }
        ($Path -match "^\\\\\?\\[a-zA-Z]{1}:\\") {
            $Path = $Path -replace [regex]::Escape("\\?\"), ""
            $result.Add($Path)
            $result = $result -replace [Regex]::Escape("\\?\"), ""
            break
        }
        default {
            # Perhaps don't terminate, but this is just for testing I guess
            Throw "Couldn't reset $Path"
        }
    }
    
    $result = $result | Sort

    return $result
}
Function Start-AltFolderSearch {
    Param([string]$FolderName)

    # Thanks Chris :-) www.christopherkibble.com

    # This exists, because...

    # Annoyingly, Get-ChildItem with forced output to an arry @(Get-ChildItem ...) can return an explicit
    # $null value for folders with no subfolders, causing the for loop to indefinitely iterate through
    # working dir when it reaches a null value, so ? $_ -ne $null is needed
    [System.Collections.ArrayList]$FolderList = @((Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue).FullName | Where-Object { [string]::IsNullOrEmpty($_) -eq $false })

    ForEach($Folder in $FolderList) {
        $FolderList += Start-AltFolderSearch -FolderName $Folder
    }

    return $FolderList
}

Function Check-FileSystemAccess {
    param
    (
        [string]$Path,
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    # Thanks to Patrick in Windows Admins slack

    [System.Security.Principal.WindowsIdentity]$currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    if (Test-Path $Path)
    {
        try
        {
            [System.Security.AccessControl.FileSystemSecurity]$security = (Get-Item -Path $Path -Force).GetAccessControl()
            if ($security -ne $null)
            {
                [System.Security.AccessControl.AuthorizationRuleCollection]$rules = $security.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])
                for([int]$i = 0; $i -lt $rules.Count; $i++)
                {
                    if (($currentIdentity.Groups.Contains($rules[$i].IdentityReference)) -or ($currentIdentity.User -eq $rules[$i].IdentityReference))
                    {
                        [System.Security.AccessControl.FileSystemAccessRule]$fileSystemRule = [System.Security.AccessControl.FileSystemAccessRule]$rules[$i]
                        if ($fileSystemRule.FileSystemRights.HasFlag($Rights))
                        {
                            return $true
                            break;
                        }
                    }
                }
            }
            else
            {
                return $false
            }
        }
        catch
        {
            return $false
        }
    }
    else
    {
        return $false
    }
}

Function Set-CMDrive {
    Param(
        [string]$SiteCode,
        [string]$Server,
        [string]$Path
    )

    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ConfigurationManager) -eq $null) {
        try {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
        } catch {
            Throw "Failed to import Configuration Manager module"
        }
    }

    try {
        # Connect to the site's drive if it is not already present
        If((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Server -ErrorAction Stop | Out-Null
        }
        # Set the current location to be the site code.
        Set-Location "$($SiteCode):\" -ErrorAction Stop

        # Verify given sitecode
        If((Get-CMSite -SiteCode $SiteCode).SiteCode -ne $SiteCode) { throw }

    } catch {
        If((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -ne $null) {
            Set-Location $Path
            Remove-PSDrive -Name $SiteCode -Force
        }
        Throw "Failed to create New-PSDrive with site code `"$($SiteCode)`" and server `"$($Server)`""
    }

}

If ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($env:COMPUTERNAME -ne $SCCMServer)) {
    # If user has given local path for $SourcesLocation, need to ensure
    # we don't produce false positives where a similar folder structure exists
    # on the remote machine and site server. e.g. packages let you specify local path
    # on site server
    Throw "Aborting: will not be able to determine unused folders using local path remote from site server"
}

[System.Collections.ArrayList]$AllContentObjects = @()
$Commands = @()

switch ($true) {
    (($Packages -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMPackage"
    }
    (($Applications -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMApplication"
    }
    (($Drivers -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMDriver"
    }
    (($DriverPackages -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMDriverPackage"
    }
    (($OSImages -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMOperatingSystemImage"
    }
    (($OSUpgradeImages -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMOperatingSystemInstaller"
    }
    (($BootImages -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMBootImage"
    }
    (($DeploymentPackages -eq $true) -Or ($All -eq $true)) {
        $Commands += "Get-CMSoftwareUpdateDeploymentPackage"
    }
}

# Get NetBIOS of given $SCCMServer parameter so it's similar format as $env:COMPUTERNAME used in body during folder/content object for loop
# And also for value pair in each content objects .AllPaths property (hashtable)
If ($SCCMServer -as [IPAddress]) {
    $FQDN = [System.Net.Dns]::GetHostEntry("$($SCCMServer)").HostName
}
Else {
    $FQDN = [System.Net.Dns]::GetHostByName($SCCMServer).HostName
}
$SCCMServer = $FQDN.Split(".")[0]

If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 0 -Status "Calculating number of folders" }

$AllFolders = Get-AllFolders -Path $SourcesLocation -AltFolderSearch $AltFolderSearch.IsPresent

$OriginalPath = (Get-Location).Path
Set-CMDrive -SiteCode $SiteCode -Server $SCCMServer -Path $OriginalPath

If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 33 -Status "Getting all CM content objects" }

$AllContentObjects = Get-CMContent -Commands $Commands -SCCMServer $SCCMServer

Set-Location $OriginalPath

$Result = @()

$AllFolders | ForEach-Object -Begin {

    If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 66 -Status "Determinig unused folders" }
    
    $NumOfFolders = $AllFolders.count

    # Forcing int data type because double/float for benefit of modulo write-progoress
    If ($NumOfFolders -ge 150) { [int]$FolderInterval = $NumOfFolders * 0.01 } else { $FolderInterval = 2 }
    $FolderCounter = 0

} -Process {

    If (($FolderCounter % $FolderInterval) -eq 0) { 
        [int]$Percentage = ($FolderCounter / $NumOfFolders * 100)
        If ($NoProgress -eq $false ) { Write-Progress -Id 2 -Activity "Looping through folders in $($SourcesLocation)" -PercentComplete $Percentage -Status "$($Percentage)% complete" -ParentId 1 }
        Write-Host "$(Get-Date): $($Percentage)%"
    }
    
    $FolderCounter++
    $Folder = $_

    $obj = New-Object PSCustomObject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Folder -Value $Folder

    $UsedBy = @()
    $AccessDenied = @()
    $IntermediatePath = $false
    $ToSkip = $false
    $NotUsed = $false

    If ((Check-FileSystemAccess -Path $Folder -Rights Read) -ne $true) {
        $UsedBy += "Access denied"
        # Still continue anyway because we can still determine if it's an exact match or intermediate path of a content object
    }

    If ($Folder.StartsWith($ToSkip)) {
        # Should probably rename $NotUsed to something more appropriate to truely reflect its meaning
        # This is here so we don't walk through completely unused folder + sub folders
        # Unused folders + sub folders are learnt for each loop of a new folder structure and thus each loop of all content objects
        $NotUsed = $true
    }
    Else {

        [int]$ContentInterval = $AllContentObjects.count * 0.25
        $ContentCounter = 0

        ForEach ($ContentObject in $AllContentObjects) {

            If ($ContentCounter % $ContentInterval -eq 0) {
                If ($NoProgress -eq $false) { Write-Progress -Id 3 -Activity "Looping through content objects" -PercentComplete ($ContentCounter / $AllContentObjects.count * 100) -ParentId 2 }
            }

            $ContentCounter++
            
            # Whatever you do, ignore case!

            switch($true) {
                ([string]::IsNullOrEmpty($ContentObject.SourcePath) -eq $true) {
                    break
                }
                ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($ContentObject.AllPaths.($Folder) -eq $env:COMPUTERNAME)) {
                    # Package is local host to the site server
                    $UsedBy += $ContentObject.Name
                    break
                }
                (($ContentObject.AllPaths.Keys -contains $Folder) -eq $true) {
                    # By default the ContainsKey method ignores case
                    $UsedBy += $ContentObject.Name
                    break
                }
                (($ContentObject.AllPaths.Keys -match [Regex]::Escape($Folder)).Count -gt 0) {
                    # If any of the content object paths start with $Folder
                    $IntermediatePath = $true
                    break
                }
                ($ContentObject.AllPaths.Keys.Where{$Folder.StartsWith($_, "CurrentCultureIgnoreCase")}.Count -gt 0) {
                    # If $Folder starts wtih any of the content object paths
                    $IntermediatePath = $true
                    break
                }
                default {
                    $ToSkip = $Folder
                    $NotUsed = $true
                }
            }

        }

        switch ($true) {
            ($UsedBy.count -gt 0) {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name UsedBy -Value (($UsedBy) -join ', ')
                # Commented out the below because if we move a found content object, it removes other relevant paths associated with it being identified as an intermediate path
                #ForEach ($item in $UsedBy) {
                #   $AllContentObjects.Remove($item) # Stop me walking through content objects that I've already found 
                #}
                break
            }
            ($IntermediatePath -eq $true) {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name UsedBy -Value "An intermediate folder (sub or parent folder)"
                break
            }
            ($NotUsed -eq $true) {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name UsedBy -Value "Not used"
                break
            }
        }
    
        $Result += $obj

    }
} -End {

    Write-Host "$(Get-Date): Complete"
    return $Result

}