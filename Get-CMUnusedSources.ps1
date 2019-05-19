<#
.SYNOPSIS
    Clear out your Configuration Manager clutter of no longer used source files on disk!

.DESCRIPTION
    This scripts allows you to identify folders that are not referenced by any content objects in your Configuration Manager environment.

.INPUTS
	Provide the script the absolute path to the "root" of your sources folder/drive

.EXAMPLE
	.\Get-AppPackageCleanUp.ps1 -PackageType Packages -SiteServer YOURSITESERVER

	.\Get-AppPackageCleanUp.ps1 -PackageType Packages -SiteServer YOURSITESERVER -ExportCSV True

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
        - Dynamic write progress (no longer possible with new .net enumeratedirectories though :())
        - Dashimo?
        - How to handle access denied on folders?
        - Get-AllLocalPath (or whatever I called it) and Get-AllSharedFolders are similar, get one to use the other or just elimnate one?
        - Test folder structures e.g. F:\Sources\More\Folders\Applications 
        - Test if a content object source path has multiple shares that are applicable to it, e.g. Applications$ and Packages$ point to F:\Sources or something like that
        - Strip trailing \ from $AllPaths
        
    Problems:
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

# Invaluable resource for getting all source locations: https://www.verboon.info/2013/07/configmgr-2012-script-to-retrieve-source-path-locations/

Function Get-CMContent {
    Param($Commands)
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
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath ($DeploymentType.Installer.Contents.Content.Location).TrimEnd('\')
                        $GetAllPathsResult = Get-AllPaths -Path $obj.SourcePath -Cache $ShareCache
                        $ShareCache = $GetAllPathsResult[0]
                        $AllPaths = $GetAllPathsResult[1]
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                        $AllContent += $obj
                    }
                }
                "Get-CMDriver" {
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.CI_ID
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.LocalizedDisplayName
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath ($item.ContentSourcePath).TrimEnd('\')
                    $GetAllPathsResult = Get-AllPaths -Path $obj.SourcePath -Cache $ShareCache
                    $ShareCache = $GetAllPathsResult[0]
                    $AllPaths = $GetAllPathsResult[1]
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                    $AllContent += $obj
                }
                default {
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.PackageId
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.Name
                    # OS images and boot iamges are absolute paths to files
                    If ("OperatingSystemImage","BootImage" -contains $obj.ContentType) {
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath (Split-Path $item.PkgSourcePath).TrimEnd('\')
                    }
                    Else {
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath ($item.PkgSourcePath).TrimEnd('\')
                    }
                    $GetAllPathsResult = Get-AllPaths -Path $obj.SourcePath -Cache $ShareCache
                    $ShareCache = $GetAllPathsResult[0]
                    $AllPaths = $GetAllPathsResult[1]
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                    $AllContent += $obj
                }   
            }
        }
    }
    #$ShareCache
    return $AllContent
}

Function Get-AllPaths {
    param (
        [string]$Path,
        [hashtable]$Cache
    )

    $AllPaths = @{}

    If ([bool]([System.Uri]$Path).IsUnc -eq $true) {
        
        # Grab server, share and remainder of path in to 4 groups: group 0 (whole match) "\\server\share\folder\folder", group 1 = "server", group 2 = "share", group 3 = "folder\folder"
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

        If ($Server -as [IPAddress]) {
            $FQDN = [System.Net.Dns]::GetHostEntry("$($Server)").HostName
            $IP = $Server
        }
        Else {
            $FQDN = [System.Net.Dns]::GetHostByName($Server).HostName
            $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
        }
        $NetBIOS = $FQDN.Split(".")[0]
        
    
        If ($Cache.ContainsKey($Server)) {
            # Already have this server's shares cached
        }
        Else {
            # Do not yet have this server's shares cached
            $Cache.Add($NetBIOS, (Get-AllSharedFolders -Server $Server))
            $Cache.Add($FQDN, (Get-AllSharedFolders -Server $Server))
            $Cache.Add($IP, (Get-AllSharedFolders -Server $Server))
        }
    
        # Verify if using UNC drive letter: match only \\server\c$
        $Regex = "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z]\$" 
        # We need to take a different approach to building $AllPaths if this is the content object's source path
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
    
            If ($Cache.$Server.ContainsKey($ShareName)) {
                $LocalPath = $Cache.$Server.$ShareName
            }
            # But, what to do if it isn't? ^
    
            $AllPathsArr = @()
            $AllPathsArr += "\\$($FQDN)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += "\\$($NetBIOS)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += "\\$($IP)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += ("\\$($FQDN)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
            $AllPathsArr += ("\\$($NetBIOS)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
            $AllPathsArr += ("\\$($IP)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
            $AllPathsArr += "$($LocalPath)$($ShareNameRemainder)"
    
            ForEach ($item in $AllPathsArr) {
                If ($AllPaths.ContainsKey($item) -eq $false) {
                    $AllPaths.Add($item, $NetBIOS)
                }
            }
    
        }
    }
    Else {
        $AllPaths.Add($Path, $env:COMPUTERNAME)
    }
    
    $result = @()
    $result += $Cache
    $result += $AllPaths
    return $result
}

Function Get-LocalPathFromUNCShare {
    param (
        [ValidatePattern("\\\\(.+)(\\).+")]
        [Parameter(Mandatory=$true)]
        [string]$Share
    )

    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)"
    $RegexMatch = [regex]::Match($Share, $Regex)
    $Server = $RegexMatch.Groups[1].Value
    $ShareName = $RegexMatch.Groups[2].Value
    
    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }

    return ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="
}

Function Get-AllSharedFolders {
    Param([String]$Server)
    # Get all shares on server
    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }
    # Iterate through them using hidden PSObject property because $Shares is PSCustomObject
    $AllShares = @{}
    $Shares.PSObject.Properties | Where-Object { $_.TypeNameOfValue -eq "Deserialized.System.String[]" } | ForEach-Object {
        # At this point it's an array
        ForEach ($item in $_) {
            $AllSharesShareName = (($item.Value -match "ShareName") -replace "ShareName=")[0] # There's only ever 1 element in the array
            $AllSharesPath = (($item.Value -match "Path") -replace "Path=")[0] # There's only ever 1 element in the array
            $AllShares += @{ $AllSharesShareName = $AllSharesPath }
        } 
    }
    return $AllShares
}

Function Set-CMDrive {
    #Try getting the site code from the client installed on this system.
    If (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\Identification" "Site Code"){
        $SiteCode =  Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\Identification" | Select-Object -ExpandProperty "Site Code"
    } ElseIf (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client" "AssignedSiteCode") {
        $SiteCode =  Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client" | Select-Object -ExpandProperty "AssignedSiteCode"
    }

    #If the client isn't installed try looking for the site code based on the PS drives.
    If (-Not ($SiteCode) ) {
        #See if a PSDrive exists with the CMSite provider
        $PSDrive = Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue

        #If PSDrive exists then get the site code from it.
        If ($PSDrive.Count -eq 1) {
            $SiteCode = $PSDrive.Name
        }
    }
    $configManagerCmdLetpath = Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
    Import-Module $configManagerCmdLetpath -Force
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
                (([bool]([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($ContentObject.AllPaths.($Folder) -eq $env:COMPUTERNAME)) {
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