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

#>
Param (
    # Sources location
    [Parameter(Mandatory=$true, Position = 1)]
    [ValidateScript({
        If (!($_ | Test-Path)) {
            Throw "Not a valid path."
        } ElseIf (!($_ | Test-Path -PathType Container)) {
            Throw "Parameter value must be a folder, not file."
        } Else {
            return $true
        }
    })]
    [string]$SourcesLocation
)

# Invaluable resource for getting all source locations: https://www.verboon.info/2013/07/configmgr-2012-script-to-retrieve-source-path-locations/

Function Get-Packages {
    Param($MasterObject)
    ForEach ($Package in (Get-CMPackage)) { 
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Package"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $Package.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $Package.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $Package.PkgSourcePath
        If (([bool]([System.Uri]$Package.PkgSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $Package.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value ($Package.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $Package.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-Drivers {
    Param($MasterObject)
    ForEach ($Driver in (Get-CMDriver)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Driver"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $Driver.CI_ID
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $Driver.LocalizedDisplayName
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $Driver.ContentSourcePath
        If (([bool]([System.Uri]$Driver.ContentSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $Driver.ContentSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value ($Driver.ContentSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $Driver.ContentSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-DriverPackages {
    Param($MasterObject)
    ForEach ($DriverPackage in (Get-CMDriverPackage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Driver Package"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $DriverPackage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $DriverPackage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $DriverPackage.PkgSourcePath
        If (([bool]([System.Uri]$DriverPackage.PkgSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $DriverPackage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value ($DriverPackage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $DriverPackage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-BootImages {
    Param($MasterObject)
    ForEach ($BootImage in (Get-CMBootImage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Boot Image"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $BootImage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $BootImage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $BootImage.PkgSourcePath
        If (([bool]([System.Uri]$BootImage.PkgSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $BootImage.PkgSourcePath
            # Using Get-Item beacuse the value is absolute path to .wim
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value (Get-Item (($BootImage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath))).DirectoryName
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $BootImage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-OSImages {
    Param($MasterObject)
    ForEach ($OSImage in (Get-CMOperatingSystemImage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Operating System Image"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $OSImage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $OSImage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $OSImage.PkgSourcePath
        If (([bool]([System.Uri]$OSImage.PkgSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $OSImage.PkgSourcePath
            # Using Get-Item beacuse the value is absolute path to .wim
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value (Get-Item (($OSImage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath))).DirectoryName
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $OSImage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-OSUpgradeImage {
    Param($MasterObject)
    ForEach ($OSUpgradeImage in (Get-CMOperatingSystemInstaller)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Operating System Upgrade Image"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $OSUpgradeImage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $OSUpgradeImage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $OSUpgradeImage.PkgSourcePath
        If (([bool]([System.Uri]$OSUpgradeImage.PkgSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $OSUpgradeImage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value ($OSUpgradeImage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $OSUpgradeImage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-DeploymentPackages {
    Param($MasterObject)
    ForEach ($DeploymentPackage in (Get-CMSoftwareUpdateDeploymentPackage)) {
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Deployment Package"
        Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $DeploymentPackage.PackageId
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $DeploymentPackage.Name
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $DeploymentPackage.PkgSourcePath
        If (([bool]([System.Uri]$DeploymentPackage.PkgSourcePath).IsUnc) -eq $true) {
            $Share = Get-LocalPathFromSharePath -Share $DeploymentPackage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value ($DeploymentPackage.PkgSourcePath -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
        }
        Else {
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $DeploymentPackage.PkgSourcePath
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
        }
        $MasterObject += $obj
    }
    return $MasterObject
}

Function Get-Applications {
    Param($MasterObject)
    ForEach ($Application in (Get-CMApplication)) {
        $AppMgmt = ([xml]$Application.SDMPackageXML).AppMgmtDigest
        $AppName = $AppMgmt.Application.DisplayInfo.FirstChild.Title
        ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Application"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value "$($DeploymentType.AuthoringScopeId)/$($DeploymentType.LogicalName)/$($DeploymentType.Version)"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value "$($Application.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $DeploymentType.Installer.Contents.Content.Location
            If (([bool]([System.Uri]$DeploymentType.Installer.Contents.Content.Location).IsUnc) -eq $true) {
                $Share = Get-LocalPathFromSharePath -Share $DeploymentType.Installer.Contents.Content.Location
                Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value ($DeploymentType.Installer.Contents.Content.Location -replace "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)", $Share.LocalPath)
                Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value ($Share.LocalPath)
            }
            Else {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePathLocal -Value $DeploymentType.Installer.Contents.Content.Location
                Add-Member -InputObject $obj -MemberType NoteProperty -Name SharePathLocal -Value $null
            }
            $MasterObject += $obj
        }
    }
    return $MasterObject
}

Function Get-AllContent {
    Clear-Variable Content -ErrorAction SilentlyContinue
    $Content = @()
    Get-Packages -MasterObject $Content
    Get-Drivers -MasterObject $Content
    Get-DriverPackages -MasterObject $Content
    Get-BootImages -MasterObject $Content
    Get-OSImages -MasterObject $Content
    Get-OSUpgradeImage -MasterObject $Content
    Get-DeploymentPackages -MasterObject $Content
    Get-Applications -MasterObject $Content
    return $Content
}

function Test-UNC {
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

Function Get-SiteCode {
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

    Return $SiteCode
}    

function Get-LocalPathFromSharePath {
    # Yet another Cody special
	param (
		[ValidatePattern("\\\\(.+)(\\).+")]
		[string]$Share
	)
	$Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)"
	$RegexMatch = [regex]::Match($Share, $Regex)
	$Server = $RegexMatch.Groups[1].Value
	$ShareName = $RegexMatch.Groups[2].Value
	
	#$Shares = Invoke-Command -ComputerName $Server -ScriptBlock {
	#	get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares
	#}
    $Shares = get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares
	$ShareLocalPath = ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="
	[PSCustomObject]@{
		ComputerName    = $Server
		LocalPath       = $ShareLocalPath
	}
}

$configManagerCmdLetpath = Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
Import-Module $configManagerCmdLetpath -Force
$SiteCode = Get-SiteCode
#If the PS drive doesn't exist then try to create it.
If (! (Test-Path "$($SiteCode):")) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root "." -WhatIf:$False | Out-Null
}
Set-Location "$($SiteCode):" | Out-Null

Write-Verbose "Getting all child folders under $SourcesLocation"
[System.Collections.ArrayList]$AllFolders = (Get-ChildItem -Directory -Recurse -Path $SourcesLocation).FullName

$AllContent = Get-AllContent 
$temp = New-Object System.Collections.ArrayList

ForEach ($folder in $AllFolders) {
    ForEach ($item in $AllContent) {
        If ($folder.StartsWith($item.SourcePathLocal)) {
            $temp.Add($folder) | Out-Null
        }
    }
}