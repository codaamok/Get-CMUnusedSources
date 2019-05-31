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
        Cody
        Chris Kibble
        Chris Dent
        PsychoData (the regex mancer)
    
    Version history:
    1

    TODO: 
        - Comment regex
        - Review comments
        - Dashimo?
        - How to handle and display access denied on folders
        - Test if a content object source path has multiple shares that are applicable to it, e.g. Applications$ and Packages$ point to F:\Sources or something like that
        - Consider using \\?\UNC\server\share format, allows you to avoid low path limit
        - Add \\?\UNC\server\share support to AllPaths property. Is this needed if Get-ChildItem fullname includes \\?\UNC\...?
        
    Problems:
        - Have I stupidly assumed share name is same as folder name on disk???
        - [RESOLVED - untested] Some content objects are absolute references to files, e.g. BootImage and OperatingSystemImage
        - [RESOLVED] Need to add local path to $AllPaths surely?
        - [RESOLVED - untested] Adding server property to $AllPaths for the below purpose. : If content object SourcePath is e.g. \\FILESERVER\SCCMSources\Applications\7zip\x64 and local path resolves to F:\SCCMSources\Applications\7zip\x64 and user gigves -SourcesLocations as F:\ and F:\Applications\7zip\x64 exists on primary site server (where script should be running from) this will produce a false positive
#>
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
            # Path that is just drive letter
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

    If ($Cache.ContainsKey($Server) -eq $false) {
        # Do not yet have this server's shares cached
        # $AllSharedFolders is null if couldn't connect to serverr to get all shared folders
        $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object {
            $AllSharedFolders = Get-AllSharedFolders -Server $Server
            If ([string]::IsNullOrEmpty($AllSharedFolders) -eq $false) {
                $Cache.Add($_, $AllSharedFolders)
            }
            Else {
                Write-Warning "Could not update cache because could not get shared folders from: `"$($Server)`" / `"$($_)`", used by `"$($obj.Name)`""
            }
        }
    }

    ##### Build the AllPaths property

    $AllPathsArr = @()

    $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object -Process {
        If ($Cache.$_.ContainsKey($ShareName)) {
            $LocalPath = $Cache.$_.$ShareName
            $AllPathsArr += ("\\$($_)\$($LocalPath)$($ShareRemainder)" -replace ':', '$')
        }
        Else {
            Write-Warning "Share `"$($ShareName)`" does not exist on `"$($_)`", used by `"$($obj.Name)`""
        }
        $AllPathsArr += "\\$($_)\$($ShareName)$($ShareRemainder)"
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
        If ($AllPaths.ContainsKey($item) -eq $false) {
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
            $AllShares += @{ $Share.Name = $Share.Path.TrimEnd("\") }
        }
    }
    catch {
        $AllShares = $null
    }

    return $AllShares
}

Function Get-AllFolders {
    # Thanks Chris :-) www.christopherkibble.com
    Param(
        [string]$FolderName
    )

    # This exists, because...

    # Annoyingly, Get-ChildItem with forced output to an arry @(Get-ChildItem ...) can return an explicit
    # $null value for folders with no subfolders, causing the for loop to indefinitely iterate through
    # working dir when it reaches a null value, so ? $_ -ne $null is needed
    [System.Collections.ArrayList]$FolderList = @((Get-ChildItem -Path $FolderName -Directory).FullName | Where-Object { [string]::IsNullOrEmpty($_) -eq $false })

    ForEach($Folder in $FolderList) {
        $FolderList += Get-AllFolders -FolderName $Folder
    }

    return $FolderList
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

# Add backslash to $SourcesLocation if given local drive letter and it's missing
# Otherwise EnumerateDirectories method happily walks through all folders without it and skews strings, e.g. F:Path\To\Folders instead of F:\Path\To\Folders
If ($SourcesLocation -match "^[a-zA-Z]:$") { $SourcesLocation = $SourcesLocation + "\" }

If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 0 -Status "Calculating number of folders" }

If ($AltFolderSearch) {
    [System.Collections.ArrayList]$AllFolders = Get-AllFolders -FolderName $SourcesLocation
}
Else {
    try {
        [System.Collections.ArrayList]$AllFolders = (Get-ChildItem -Path $SourcesLocation -Directory -Recurse).FullName
    }
    catch {
        Throw "Consider using -AltFolderSearch"
    }
}

$AllFolders.Add($SourcesLocation)
$AllFolders = $AllFolders | Sort

$OriginalPath = (Get-Location).Path
Set-CMDrive -SiteCode $SiteCode -Server $SCCMServer -Path $OriginalPath

If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 33 -Status "Getting all CM content objects" }
$AllContentObjects = Get-CMContent -Commands $Commands -SCCMServer $SCCMServer

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
    $IntermediatePath = $false
    $ToSkip = $false
    $NotUsed = $false

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
                    # Package is local host
                    $UsedBy += $ContentObject
                    break
                }
                ($ContentObject.AllPaths.ContainsKey($Folder) -eq $true) {
                    # By default the ContainsKey method ignores case
                    $UsedBy += $ContentObject
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
                Add-Member -InputObject $obj -MemberType NoteProperty -Name UsedBy -Value (($UsedBy.Name) -join ', ')
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

    Write-Host "$(Get-Date): 100%"

    Set-Location $OriginalPath
    
    # return $Result

}