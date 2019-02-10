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
    [string]$Folder
)

$AllFolders = (Get-ChildItem -Directory -Recurse).FullName