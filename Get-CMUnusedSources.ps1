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
    
    Version history:
    1

    TODO:
        - Comment regex
        - Dashimo?
        - How to handle access denied on folders?
        
    Problems:
        - Some content objects are absolute references to files, e.g. boot images and either OS or OS upgrade images
        - Need to add local path to $AllPaths surely? Perhaps need a "Server" property too in case of: 
        - Have I stupidly assumed share name is same as folder name on disk???

    Think I fixed? Not 100% confident:
        - Adding server property to $AllPaths for the below purpose. : If content object SourcePath is e.g. \\FILESERVER\SCCMSources\Applications\7zip\x64 and local path resolves to F:\SCCMSources\Applications\7zip\x64 and user gigves -SourcesLocations as F:\ and F:\Applications\7zip\x64 exists on primary site server (where script should be running from) this will produce a false positive


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
        If (!($_ | Test-Path)) {
            Throw "Invalid path or insufficient permissions"
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

Function Get-AllContentMinusApplicationsAndDrivers {
    $Commands = "Get-CMPackage", "Get-CMDriverPackage", "Get-CMBootImage", "Get-CMOperatingSystemImage", "Get-CMOperatingSystemInstaller", "Get-CMSoftwareUpdateDeploymentPackage"
    $MasterObject = @()
    $ShareCache = @{}
    ForEach ($Command in $Commands) {
        ForEach ($item in (Invoke-Expression $Command)) {
            $AllPaths = @{}
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
            Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.PackageId
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.Name
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $item.PkgSourcePath
            If ([bool]([System.Uri]$item.PkgSourcePath).IsUnc -eq $true) {
                # Check cache
                $Regex = "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]*"
                $RegexMatch = [regex]::Match($item.PkgSourcePath,$Regex)
                If ($ShareCache.ContainsKey($RegexMatch.Groups[0].Value) -eq $false) {
                    $ShareLocalPath = Get-LocalPathFromUNCShare -Share $item.PkgSourcePath
                    $ShareCache.Add($RegexMatch.Groups[0].Value, $ShareLocalPath)
                }
                Else {
                    $ShareLocalPath = $ShareCache.($RegexMatch.Groups[0].Value)
                }
                $AllPaths = Get-AllPossibleUNCPaths -Share $item.PkgSourcePath -LocalPath $ShareLocalPath
                Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                # Update cache
                ForEach ($Path in $AllPaths.GetEnumerator()) {
                    If ([bool]([System.Uri]$Path.Name).IsUnc -eq $true) { # Don't add the local path to the cache
                        $Regex = "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]*" # grab Server + Share, e.g. \\server\share
                        $RegexMatch = [regex]::Match($Path.Name,$Regex)
                        $Regex = "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z]\$" # grab Server + drive share, e.g. \\server\c$
                        If (($ShareCache.ContainsKey($RegexMatch.Groups[0].Value) -eq $false) -And ($RegexMatch.Groups[0].Value -notmatch $Regex)) {
                            $ShareCache.Add($RegexMatch.Groups[0].Value,$ShareLocalPath)
                        }
                    }
                }
            }
            Else {
                $AllPaths.Add($item.PkgSourcePath, $env:COMPUTERNAME)
                Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
            }
            $MasterObject += $obj
        }
    }
    $ShareCache
    $MasterObject
}

 # Drivers and Applications on their own

Function Get-Drivers {
    $MasterObject = @()
    ForEach ($Driver in (Get-CMDriver)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Driver"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $Driver.CI_ID
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $Driver.LocalizedDisplayName
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $Driver.ContentSourcePath
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-Applications {
    $MasterObject = @()
    ForEach ($Application in (Get-CMApplication)) {
        $AppMgmt = ([xml]$Application.SDMPackageXML).AppMgmtDigest
        $AppName = $AppMgmt.Application.DisplayInfo.FirstChild.Title
        ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Application"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value "$($DeploymentType.AuthoringScopeId)/$($DeploymentType.LogicalName)"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value "$($Application.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $DeploymentType.Installer.Contents.Content.Location
            $MasterObject += $obj
        }
    }
    return $MasterObject
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

Function Get-AllPossibleUNCPaths {
    param (
        [ValidatePattern("\\\\(.+)(\\).+")]
        [Parameter(Mandatory=$true)]
        [string]$Share,
        [string]$LocalPath
    )
    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~\\!@#$%^&(){}\'._-]+)*"
    $RegexMatch = [regex]::Match($Share, $Regex)

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

    $FQDN = [System.Net.Dns]::GetHostByName(("$($Server)")).HostName
    $NetBIOS = $FQDN.Split(".")[0]
    $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString

    $result = @{}

    $result.Add("\\$($FQDN)\$($ShareName)$($ShareNameRemainder)", $NetBIOS)
    $result.Add(("\\$($FQDN)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$'), $NetBIOS)
    $result.Add("\\$($NetBIOS)\$($ShareName)$($ShareNameRemainder)", $NetBIOS)
    $result.Add(("\\$($NetBIOS)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$'), $NetBIOS)
    $result.Add("\\$($IP)\$($ShareName)$($ShareNameRemainder)", $NetBIOS)
    $result.Add(("\\$($IP)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$'), $NetBIOS)
    $result.Add("$($LocalPath)$($ShareNameRemainder)",$NetBIOS)

    return $result
}

Function Get-LocalPathFromUNCShare {
    param (
        [ValidatePattern("\\\\(.+)(\\).+")]
        [Parameter(Mandatory=$true)]
        [string]$Share
    )

    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)"
    $RegexMatch = [regex]::Match($item.PkgSourcePath, $Regex)
    $Server = $RegexMatch.Groups[1].Value
    $ShareName = $RegexMatch.Groups[2].Value
    
    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }

    return ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="
}

$OriginalPath = (Get-Location).Path
Set-CMDrive

[System.Collections.ArrayList]$AllContentObjects = @()

switch ($true) {
    (($Packages -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-Packages
    }
    (($Applications -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-Applications
    }
    (($Drivers -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-Drivers
    }
    (($DriverPackages -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-DriverPackages
    }
    (($OSImages -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-OSImages
    }
    (($OSUpgradeImages -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-OSUpgradeImage
    }
    (($BootImages -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-BootImages
    }
    (($DeploymentPackages -eq $true) -Or ($All -eq $true)) {
        $AllContentObjects += Get-DeploymentPackages
    }
}

# Must be a beter way than this
[System.Collections.ArrayList]$AllFolders = (Get-ChildItem -Directory -Recurse -Path $SourcesLocation).FullName
$AllFolders.Add($SourcesLocation) # Add what the user gave us
$AllFolders = $AllFolders | Sort

$Shares = @{}

ForEach ($Folder in $AllFolders) { # For every folder

    ForEach ($ContentObject in $AllContentObjects) { # For every content object

        If ([bool]([System.Uri]$SourcesLocation).IsUnc -eq $true) {

        } Else {

        }

    }

}

Set-Location $OriginalPath