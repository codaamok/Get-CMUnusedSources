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
        - Dashimo?
        - Get-AllLocalPath (or whatever I called it) and Get-AllSharedFolders are similar, get one to use the other or just elimnate one?
        - Test if a content object source path has multiple shares that are applicable to it, e.g. Applications$ and Packages$ point to F:\Sources or something like that
        - optionally show progress, faster without
        - Adjust to run from any machine
        - As a result of the last machine, remove need for UAC and add mandatory parameters specifying servername + site code?
        
    Problems:
        - What if a content object source path is \\server\share ?
        - What if content object has no path specified?
        - Have I stupidly assumed share name is same as folder name on disk???
        - [RESOLVED - untested] Some content objects are absolute references to files, e.g. BootImage and OperatingSystemImage
        - [RESOLVED] Need to add local path to $AllPaths surely?
        - [RESOLVED - untested] Adding server property to $AllPaths for the below purpose. : If content object SourcePath is e.g. \\FILESERVER\SCCMSources\Applications\7zip\x64 and local path resolves to F:\SCCMSources\Applications\7zip\x64 and user gigves -SourcesLocations as F:\ and F:\Applications\7zip\x64 exists on primary site server (where script should be running from) this will produce a false positive
#>
#Requires -RunAsAdministrator
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
    [switch]$DeploymentPackages
)

Function Get-CMContent {
    Param($Commands)
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
                        $obj = New-Object PSObject
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value "$($DeploymentType.AuthoringScopeId)/$($DeploymentType.LogicalName)"
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value "$($item.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
                        If ($DeploymentType.Installer.Contents.Content.Location -ne $null) {
                            $SourcePath = ($DeploymentType.Installer.Contents.Content.Location).TrimEnd('\')
                            $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SCCMServer $SCCMServer
                            $ShareCache = $GetAllPathsResult[0]
                            $AllPaths = $GetAllPathsResult[1]
                        }
                        Else {
                            $SourcePath = $null
                            $AllPaths = $null
                        }
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                        $AllContent += $obj
                    }
                }
                "Get-CMDriver" { # I don't actually think it's possible for a driver to not have source path set
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.CI_ID
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.LocalizedDisplayName
                    If ($item.ContentSourcePath -ne $null) {
                        $SourcePath = ($item.ContentSourcePath).TrimEnd('\')
                        $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SCCMServer $SCCMServer
                        $ShareCache = $GetAllPathsResult[0]
                        $AllPaths = $GetAllPathsResult[1]
                    }
                    Else {
                        $SourcePath = $null
                        $AllPaths = $null
                    }
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                    $AllContent += $obj
                }
                default {
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.PackageId
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.Name
                    If ($item.PkgSourcePath -ne $null) {
                        # OS images and boot iamges are absolute paths to files
                        If ("OperatingSystemImage","BootImage" -contains $obj.ContentType) {
                            $SourcePath = (Split-Path $item.PkgSourcePath).TrimEnd('\')
                        }
                        Else {
                            $SourcePath = ($item.PkgSourcePath).TrimEnd('\')
                        }
                        $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SCCMServer $SCCMServer
                        $ShareCache = $GetAllPathsResult[0]
                        $AllPaths = $GetAllPathsResult[1]
                    }
                    Else {
                        $SourcePath = $null
                        $AllPaths = $null
                    }
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                    $AllContent += $obj
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

    If ([bool]([System.Uri]$Path).IsUnc -eq $true) {
        
        ############################################
        # Grab server, share and remainder of path#in to 4 groups:
        # group 0 (whole match) "\\server\share\folder\folder", group 1 = "server", group 2 = "share", group 3 = "folder\folder"
        ############################################

        $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)\\([a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)*" 
        $RegexMatch = [regex]::Match($Path, $Regex)
        switch ($true) {
            ($RegexMatch.Groups[1].Success -eq $true) {
                $Server = $RegexMatch.Groups[1].Value
            }
            ($RegexMatch.Groups[2].Success -eq $true) {
                $ShareName = $RegexMatch.Groups[2].Value
            }
            ($RegexMatch.Groups[3].Success -eq $true) {
                $ShareNameRemainder = "\" + $RegexMatch.Groups[3].Value
            }
            default { # do some sort of error handling with this later? prob not necessary as .IsUnc from its caller probably qualifiees it already?
                $Server = ""
                $ShareName = ""
                $ShareNameRemainder = ""
            }
        }

        ############################################
        # Determine FQDN, IP and NetBIOS
        ############################################

        If ($Server -as [IPAddress]) {
            $FQDN = [System.Net.Dns]::GetHostEntry("$($Server)").HostName
            $IP = $Server
        }
        Else {
            $FQDN = [System.Net.Dns]::GetHostByName($Server).HostName
            $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
        }
        $NetBIOS = $FQDN.Split(".")[0]
        
        ############################################
        # Update the cache of shared folders and their local paths
        ############################################
    
        If ($Cache.ContainsKey($Server)) {
            # Already have this server's shares cached
        }
        Else {
            # Do not yet have this server's shares cached
            # $AllSharedFolders is null if couldn't connect to serverr to get all shared folders
            $NetBIOS,$FQDN,$IP | ForEach-Object {
                $AllSharedFolders = Get-AllSharedFolders -Server $Server
                If ($AllSharedFolders -ne $null) {
                    $Cache.Add($_, $AllSharedFolders)
                }
            }
        }

        ############################################
        # Build the AllPaths property
        ############################################
    
        # Verify if the path is using drive letter, e.g. match only \\server\c$
        # A different approach is taken to determine AllPaths if the path uses driver letter vs not a driver letter
        $Regex = "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z]\$" 
        If ($Path -match $Regex) {
            # Convert UNC path to local path
            $LocalPath = $Path -replace "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\" 
            $LocalPath = $LocalPath -replace "\$",":"
            # There was a need to capitalise the driver letter but I can't remember why now
            $LocalPath = [regex]::replace($LocalPath, "^[a-z]:\\", { $args[0].Value.ToUpper() })
    
            # Start building $AllPaths with what we already know
            $AllPaths.Add($LocalPath, $NetBIOS)
            
            # Now determine all possible paths
            If ($Cache.$Server.count -ge 1) {
                ForEach ($Share in $Cache.$Server.GetEnumerator()) {
                    If($LocalPath.StartsWith($Share.Value, "CurrentCultureIgnoreCase")) {
                        $AllPathsArr = @()
                        $AllPathsArr += $LocalPath.replace($Share.Value, "\\$($FQDN)\$($Share.Name)")
                        $AllPathsArr += $LocalPath.replace($Share.Value, "\\$($NetBIOS)\$($Share.Name)")
                        $AllPathsArr += $LocalPath.replace($Share.Value, "\\$($IP)\$($Share.Name)")
                        $AllPathsArr += (("\\" + $FQDN + "\" + $LocalPath) -replace ":", "$")
                        $AllPathsArr += (("\\" + $NetBIOS + "\" + $LocalPath) -replace ":", "$")
                        $AllPathsArr += (("\\" + $IP + "\" + $LocalPath) -replace ":", "$")
                        # This is so we can avoid error messages about dictionary already containing key
                        # Dupes can occur if there are multiple shares within the path
                        ForEach ($item in $AllPathsArr) {
                            If ($AllPaths.ContainsKey($item) -eq $false) {
                                $AllPaths.Add($item, $NetBIOS)
                            }
                        }
                    }
                }
            }
            Else {
                $AllPaths.Add($Path, $NetBIOS)
            }
        }
        Else {

            $AllPathsArr = @()

            If ($Cache.$Server.ContainsKey($ShareName)) {
                $LocalPath = $Cache.$Server.$ShareName
                $AllPathsArr += ("\\$($FQDN)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
                $AllPathsArr += ("\\$($NetBIOS)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
                $AllPathsArr += ("\\$($IP)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
                $AllPathsArr += "$($LocalPath)$($ShareNameRemainder)"
            }

            $AllPathsArr += "\\$($FQDN)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += "\\$($NetBIOS)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += "\\$($IP)\$($ShareName)$($ShareNameRemainder)"

            ForEach ($item in $AllPathsArr) {
                If ($AllPaths.ContainsKey($item) -eq $false) {
                    $AllPaths.Add($item, $NetBIOS)
                }
            }
    
        }
    }
    Else {
        $AllPaths.Add($Path, $SCCMServer)
    }
    
    $result = @()
    $result += $Cache
    $result += $AllPaths
    return $result
}

Function Get-AllSharedFolders {
    Param([String]$Server)
    # Get all shares on server
    try {
        $Shares = Invoke-Command -ComputerName $Server -ErrorAction SilentlyContinue -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }
    }
    catch {
        $Shares = $null
    }
    # Iterate through them using hidden PSObject property because $Shares is PSCustomObject
    $AllShares = @{}
    If ($Shares -ne $null) {
        $Shares.PSObject.Properties | Where-Object { $_.TypeNameOfValue -eq "Deserialized.System.String[]" } | ForEach-Object {
            # At this point it's an array
            ForEach ($item in $_) {
                $AllSharesShareName = (($item.Value -match "ShareName") -replace "ShareName=")[0] # There's only ever 1 element in the array
                $AllSharesPath = (($item.Value -match "Path") -replace "Path=")[0] # There's only ever 1 element in the array
                $AllShares += @{ $AllSharesShareName = $AllSharesPath }
            }
        }
    }
    return $AllShares
}

Function Set-CMDrive {
    Import-Module $env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1
    #If the PS drive doesn't exist then try to create it.
    If (! (Test-Path "$($SiteCode):")) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root "." | Out-Null
    }
    Set-Location "$($SiteCode):" | Out-Null
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

#Write-Debug "Getting folders"
# Must be a beter way than this
#[System.Collections.ArrayList]$AllFolders = (Get-ChildItem -Directory -Recurse -Path $SourcesLocation).FullName
# Add what the user gave us
#$AllFolders.Add($SourcesLocation) 
#$AllFolders = $AllFolders | Sort
If ($SourcesLocation -match "^[a-zA-Z]:$") { $SourcesLocation = $SourcesLocation + "\" }

$OriginalPath = (Get-Location).Path
Set-CMDrive

#Write-Debug "Getting content"
Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 0 -Status "Getting all CM content objects"
$AllContentObjects = Get-CMContent -Commands $Commands

$Result = @()

Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 33 -Status "Calculating number of folders"
$NumOfFolders = ([System.IO.Directory]::EnumerateDirectories($SourcesLocation, "*", "AllDirectories") | Measure-Object).count
# Forcing int data type because double/float for benefit of modulo write-progoress
[int]$interval = $NumOfFolders * 0.01
$counter = 0

Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 66 -Status "Determinig unused folders"
[System.IO.Directory]::EnumerateDirectories($SourcesLocation, "*", "AllDirectories") | ForEach-Object {

    If (($counter % $interval) -eq 0) { 
        [int]$Percentage = ($counter / $NumOfFolders * 100)
        Write-Progress -Id 2 -Activity "Looping through folders in $($SourcesLocation)" -PercentComplete $Percentage -Status "$($Percentage)% complete" -ParentId 1
        Write-Host "$(Get-Date): $($Percentage)%"
    }
    
    $counter++
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

        [int]$interval2 = $AllContentObjects.count * 0.25
        $counter2 = 0

        ForEach ($ContentObject in $AllContentObjects) {

            If ($counter2 % $interval2 -eq 0) {
                Write-Progress -Id 3 -Activity "Looping through content objects" -PercentComplete ($counter2 / $AllContentObjects.count * 100) -ParentId 2
            }

            $counter2++
            
            # Whatever you do, ignore case!

            switch($true) {
                ($ContentObject.SourcePath -eq $null) {
                    break
                }
                (([bool]([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($ContentObject.AllPaths.($Folder) -eq $SCCMServer)) {
                    # Package is local host
                    # Heavily assumes this scripts runs from primary site
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
}

Set-Location $OriginalPath