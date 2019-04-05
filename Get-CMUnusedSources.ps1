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
        What if people mix up their source UNC path? e.g. \\sccm\applications$ or \\sccm.acc.local\applications$ or \\sccm\d$\Applications or \\192.168.0.10\applications$ ... etc ....
        Test-Path, ensure it's still alive?
        Dashimo?

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
        If ([bool]([System.Uri]$Package.PkgSourcePath).IsUnc -eq $false) {
            Throw "Must provide UNC path"
        } ElseIf (!($_ | Test-Path)) {
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

Function Get-Packages {
    $MasterObject = @()
    ForEach ($Package in (Get-CMPackage)) { 
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Package"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $Package.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $Package.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $Package.PkgSourcePath
        $MasterObject += $obj
    }
    return $MasterObject
}

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

Function Get-DriverPackages {
    $MasterObject = @()
    ForEach ($DriverPackage in (Get-CMDriverPackage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Driver Package"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $DriverPackage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $DriverPackage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $DriverPackage.PkgSourcePath
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-BootImages {
    $MasterObject = @()
    ForEach ($BootImage in (Get-CMBootImage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Boot Image"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $BootImage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $BootImage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath (Get-Item (($BootImage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath))).DirectoryName
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-OSImages {
    $MasterObject = @()
    ForEach ($OSImage in (Get-CMOperatingSystemImage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Operating System Image"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $OSImage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $OSImage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath (Get-Item (($OSImage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath))).DirectoryName
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-OSUpgradeImage {
    $MasterObject = @()
    ForEach ($OSUpgradeImage in (Get-CMOperatingSystemInstaller)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Operating System Upgrade Image"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $OSUpgradeImage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $OSUpgradeImage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $OSUpgradeImage.PkgSourcePath
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-DeploymentPackages {
    $MasterObject = @()
    ForEach ($DeploymentPackage in (Get-CMSoftwareUpdateDeploymentPackage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Deployment Package"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $DeploymentPackage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $DeploymentPackage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $DeploymentPackage.PkgSourcePath
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

Function Test-UNC {
    #http://blogs.microsoft.co.il/scriptfanatic/2010/05/27/quicktip-how-to-validate-a-unc-path/ 
    Param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('FullName')]
        [System.String[]]$Path
    )
    Process {
        ForEach($p in $Path) {
            [bool]([System.Uri]$p).IsUnc
        }
    }
}

Function Set-CMDrive {
    <#
    .SYNOPSIS
       Attempt to determine the current device's site code from the registry or PS drive.
       Author: Bryan Dam (damgoodadmin.com)
    
    .DESCRIPTION
       When ran this function will look for the client's site.  If not found it will look for a single PS drive.
    
    .EXAMPLE
       Get-SiteCode
    
    #>
    
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
		[string]$Share
    )
    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)"
	$RegexMatch = [regex]::Match($Share, $Regex)
	$Server = $RegexMatch.Groups[1].Value
    $ShareName = $RegexMatch.Groups[2].Value

    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }
    $ShareLocalPath = ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="

    $FQDN = [System.Net.Dns]::GetHostByName(("$($server)")).HostName
    $NetBIOS = $FQDN.Split(".")[0]
    $IP = (((Test-Connection $server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
    $Full = Get-LocalPathFromSharePath -Share $Share

    $All = @()
    $All += "\\$($NetBIOS)\$($ShareName)"
    $All += "\\$($NetBIOS)\$($ShareLocalPath -replace ':', '$')"
    $All += "\\$($FQDN)\$($ShareName)"
    $All += "\\$($FQDN)\$($ShareLocalPath -replace ':', '$')"
    $All += "\\$($IP)\$($ShareName)"
    $All += "\\$($IP)\$($ShareLocalPath -replace ':', '$')"

    return $All
}

$OriginalPath = (Get-Location).Path
Set-CMDrive

Clear-Variable AllContent -ErrorAction SilentlyContinue
[System.Collections.ArrayList]$AllContent = @()

switch ($true) {
    (($Packages -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-Packages
    }
    (($Applications -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-Applications
    }
    (($Drivers -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-Drivers
    }
    (($DriverPackages -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-DriverPackages
    }
    (($OSImages -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-OSImages
    }
    (($OSUpgradeImages -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-OSUpgradeImage
    }
    (($BootImages -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-BootImages
    }
    (($DeploymentPackages -eq $true) -Or ($All -eq $true)) {
        $AllContent += Get-DeploymentPackages
    }
}

[System.Collections.ArrayList]$AllFolders = (Get-ChildItem -Directory -Recurse -Path $SourcesLocation).FullName

$Results = @()

ForEach ($Folder in $AllFolders) { # For every folder

    Write-Progress -Activity "Looping through folders" -CurrentOperation "$Folder" -id 1 -PercentComplete (($AllFolders.IndexOf($Folder) / $AllFolders.count) * 100) -Status "Finding content that uses"
    
    $obj = New-Object PSObject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Folder -Value $Folder

    $UsedBy = @()
    $IntermediatePath = $false
    $ToSkip = $false

    If ($Folder.StartsWith($ToSkip)) { # Don't walk through folders that aren't a sub or parent folder for any content objects
        $NotUsed = $true
    }
    Else {
        ForEach ($item in $AllContent) { # For every content object
            $SourcePathLocalTrimmed = ($item.SourcePathLocal).TrimEnd("\")
            $FolderTrimmed = ($Folder).TrimEnd("\")
            switch ($true) {
                ($SourcePathLocalTrimmed -eq $FolderTrimmed) {
                    $UsedBy += $item
                    break
                }
                (($FolderTrimmed.StartsWith($SourcePathLocalTrimmed) -Or ($SourcePathLocalTrimmed.StartsWith($FolderTrimmed)))) {
                    $IntermediatePath = $true
                    break
                }
                default {
                    $ToSkip = $Folder
                    $NotUsed = $true
                }
            }
        }

    }

    switch ($true) {
        ($UsedBy.count -gt 0) {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name UsedBy -Value (($UsedBy.Name) -join ', ')
            ForEach ($item in $UsedBy) {
               $AllContent.Remove($item) # Stop me walking through content objects that I've already found 
            }
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

    $Results += $obj
    
}

Set-Location $OriginalPath

return $Results, $AllContent