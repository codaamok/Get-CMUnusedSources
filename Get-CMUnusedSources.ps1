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
    ForEach ($Package in (Get-CMPackage)) { 
        [PSCustomObject] @{
            Object = "Package"
            UniqueID = $Package.PackageId
            Name = $Package.Name
            SourcePath = $Package.PkgSourcePath
            IsUNC = [bool]([System.Uri]$Package.PkgSourcePath).IsUnc
        }
    }
}

Function Get-Drivers {
    ForEach ($Driver in (Get-CMDriver)) {
        [PSCustomObject] @{
            Object = "Driver"
            UniqueID = $Driver.CI_ID
            Name = $Driver.LocalizedDisplayName
            SourcePath = $Driver.ContentSourcePath
            IsUNC = [bool]([System.Uri]$Driver.ContentSourcePath).IsUnc
        }
    }
}

Function Get-DriverPackages {
    ForEach ($DriverPackage in (Get-CMDriverPackage)) {
        [PSCustomObject] @{
            Object = "Driver Package"
            UniqueID = $DriverPackage.PackageId
            Name = $DriverPackage.Name
            SourcePath = $DriverPackage.PkgSourcePath
            IsUNC = [bool]([System.Uri]$DriverPackage.PkgSourcePath).IsUnc
        }
    }
}

Function Get-BootImages {
    ForEach ($BootImage in (Get-CMBootImage)) {
        [PSCustomObject] @{
            Object = "Boot Image"
            UniqueID = $BootImage.PackageId
            Name = $BootImage.Name
            SourcePath = $BootImage.PkgSourcePath
            IsUNC = [bool]([System.Uri]$BootImage.PkgSourcePath).IsUnc
        }
    }
}

Function Get-OSImages {
    ForEach ($OSImage in (Get-CMOperatingSystemImage)) {
        [PSCustomObject] @{
            Object = "Operating System Image"
            UniqueID = $OSImage.PackageId
            Name = $OSImage.Name
            SourcePath = $OSImage.PkgSourcePath
            IsUNC = [bool]([System.Uri]$OSImage.PkgSourcePath).IsUnc
        }
    }
}

Function Get-OSUpgradeImage {
    ForEach ($OSUpgradeImage in (Get-CMOperatingSystemInstaller)) {
        [PSCustomObject] @{
            Object = "Operating System Upgrade Image"
            UniqueID = $OSUpgradeImage.PackageId
            Name = $OSUpgradeImage.Name
            SourcePath = $OSUpgradeImage.PkgSourcePath
            IsUNC = [bool]([System.Uri]$OSUpgradeImage.PkgSourcePath).IsUnc
        }
    }
}

Function Get-DeploymentPackages {
    ForEach ($DeploymentPackage in (Get-CMSoftwareUpdateDeploymentPackage)) {
        [PSCustomObject] @{
            Object = "Deployment Package"
            UniqueID = $DeploymentPackage.PackageId
            Name = $DeploymentPackage.Name
            SourcePath = $DeploymentPackage.PkgSourcePath
            IsUNC = [bool]([System.Uri]$DeploymentPackage.PkgSourcePath).IsUnc
        }
    }
}

Function Get-Applications {
    ForEach ($Application in (Get-CMApplication)) {
        $AppMgmt = ([xml]$Application.SDMPackageXML).AppMgmtDigest
        $AppName = $AppMgmt.Application.DisplayInfo.FirstChild.Title
        ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
            [PSCustomObject] @{
                Object = "Application"
                UniqueId = ($Application.ModelName)
                Name = "$($Application.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
                SourcePath = $DeploymentType.Installer.Contents.Content.Location
                IsUNC = [bool]([System.Uri]$DeploymentType.Installer.Contents.Content.Location).IsUnc
            }
        }
    }
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
	
	$Shares = Invoke-Command -ComputerName $Server -ScriptBlock {
		get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares
	}
	$LocalPath = ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="
	[PSCustomObject]@{
		ComputerName = $Server
		LocalPath    = $LocalPath
	}
}

Write-Verbose "Getting all child folders under $SourcesLocation"
$AllFolders = (Get-ChildItem -Directory -Recurse -Path $SourcesLocation).FullName

Write-Verbose "Getting all SMB shares on $env:computername"
$AllShares = Get-SmbShare

Function Get-AllContent {
    Get-Packages
    Get-Drivers
    Get-DriverPackages
    Get-BootImages
    Get-OSImages
    Get-OSUpgradeImage
    Get-DeploymentPackages
    Get-Applications
}