<#
.SYNOPSIS
Get-CMUnusedSources will tell you what folders are not used by ConfigMgr in a given path.

.DESCRIPTION
Check out https://www.cookadam.co.uk/get-cmunusedsources and https://github.com/codaamok/Get-CMUnusedSources.

.PARAMETER SourcesLocation
The path to the directory you store your ConfigMgr sources. Can be a UNC or local path. Must be a valid path that you have read access to.

.PARAMETER SiteCode
The site code of the ConfigMgr site you wish to query for content objects.

.PARAMETER SiteServer
The site server of the given ConfigMgr site code. The server must be reachable over a network.

.PARAMETER Packages
Specify this switch to include Packages within the search to determine unused content on disk.

.PARAMETER Applications
Specify this switch to include Applications within the search to determine unused content on disk.

.PARAMETER Drivers
Specify this switch to include Drivers within the search to determine unused content on disk.

.PARAMETER DriverPackages
Specify this switch to include DriverPackages within the search to determine unused content on disk.

.PARAMETER OSImages
Specify this switch to include OSImages within the search to determine unused content on disk.

.PARAMETER OSUpgradeImages
Specify this switch to include OSUpgradeImages within the search to determine unused content on disk.

.PARAMETER BootImages
Specify this switch to include BootImages within the search to determine unused content on disk.

.PARAMETER DeploymentPackages
Specify this switch to include DeploymentPackages within the search to determine unused content on disk.

.PARAMETER AltFolderSearch
Specify this if you suspect there are issue with the default mechanism of gathering folders, which is:
    Get-ChildItem -LiteralPath "\\?\UNC\server\share\folder" -Directory -Recurse | Select-Object -ExpandProperty FullName

.PARAMETER NoProgress
Specify this to disable use of Write-Progress.

.PARAMETER Log
Specify this to enable logging. The log file(s) will be saved to the same directory as this script with a name of <scriptname>_<datetime>.log. Rolled log files will follow a naming convention of <filename>_1.lo_ where the int increases for each rotation. Each maximum log file is 2MB.

.PARAMETER ExportReturnObject
Specify this option if you wish to export the PowerShell return object to an XML file. The XML file be saved to the same directory as this script with a name of <scriptname>_<datetime>_result.xml. It can easily be reimported using Import-Clixml cmdlet.

.PARAMETER ExportCMContentObjects
Specify this option if you wish to export all ConfigMgr content objects to an XML file. The XML file be saved to the same directory as this script with a name of <scriptname>_<datetime>_cmobjects.xml. It can easily be reimported using Import-Clixml cmdlet.

.PARAMETER HtmlReport
Specify this option to enable the generation for a HTML report of the result. Doing this will force you to have the PSWriteHtml module installed. For more information on PSWriteHTML: https://github.com/EvotecIT/PSWriteHTML. The HTML file will be saved to the same directory as this script with a name of <scriptname>_<datetime>.html.

.PARAMETER Threads
Set the number of threads you wish to use for concurrent processing of this script. Default value is number of processes from env var NUMBER_OF_PROCESSORS. 

.INPUTS

.OUTPUTS

.EXAMPLE
C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation \\sccm\Applications$ -SiteCode ACC -SiteServer SCCM -Applications -Log -LogFileSize 10MB -NumOfRotatedLogs 5 -ExportReturnObject -HtmlReport -Threads 2

.EXAMPLE
C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation F:\ -SiteCode ACC -SiteServer SCCM -Log -HtmlReport

.NOTES
Author:     Adam Cook (@codaamok)
Updated:    23/07/2019
License:    GLP-3.0
Source:     https://github.com/codaamok/Get-CMUnusedSources  
#>
#Requires -Version 5.1
Param (
    [Parameter(Mandatory=$true, Position = 0, HelpMessage="Valid path (local or remote0 to where you store you ConfigMgr sources.")]
    [ValidateScript({
        If (!([System.IO.Directory]::Exists($_))) {
            throw "Invalid path or access denied"
        } ElseIf (!($_ | Test-Path -PathType Container)) {
            throw "Value must be a directory, not a file"
        } Else {
            return $true
        }
    })]
    [string]$SourcesLocation,
    [Parameter(Mandatory=$true, Position = 1, HelpMessage="ConfigMgr site code you are querying.")]
    [ValidatePattern('^[a-zA-Z0-9]{3}$')]
    [string]$SiteCode,
    [Parameter(Mandatory=$true, Position = 2, HelpMessage="ConfigMgr site server of the site site code.")]
    [ValidateScript({
        If(!(Test-Connection -ComputerName $_ -Count 1 -ErrorAction SilentlyContinue)) {
            throw "Host `"$($_)`" is unreachable"
        } Else {
            return $true
        }
    })]
    [string]$SiteServer,
    [Parameter(Mandatory=$false, HelpMessage="Gather packages.")]
    [switch]$Packages,
    [Parameter(Mandatory=$false, HelpMessage="Gather applications.")]
    [switch]$Applications,
    [Parameter(Mandatory=$false, HelpMessage="Gather drivers.")]
    [switch]$Drivers,
    [Parameter(Mandatory=$false, HelpMessage="Gather driver packages.")]
    [switch]$DriverPackages,
    [Parameter(Mandatory=$false, HelpMessage="Gather Operating System images.")]
    [switch]$OSImages,
    [Parameter(Mandatory=$false, HelpMessage="Gather Operating System upgrade images.")]
    [switch]$OSUpgradeImages,
    [Parameter(Mandatory=$false, HelpMessage="Gather boot images.")]
    [switch]$BootImages,
    [Parameter(Mandatory=$false, HelpMessage="Gather deployment packages.")]
    [switch]$DeploymentPackages,
    [Parameter(Mandatory=$false, HelpMessage="Enable alternative folder search.")]
    [switch]$AltFolderSearch,
    [Parameter(Mandatory=$false, HelpMessage="Disable use of Write-Progress.")]
    [switch]$NoProgress,
    [Parameter(Mandatory=$false, HelpMessage="Enable logging.")]
    [switch]$Log,
    [Parameter(Mandatory=$false, HelpMessage="Generate XML export of PowerShell object with the result.")]
    [switch]$ExportReturnObject,
    [Parameter(Mandatory=$false, HelpMessage="Generate XML export of PowerShell object with all ConfigMgr content objects.")]
    [switch]$ExportCMContentObjects,
    [Parameter(Mandatory=$false, HelpMessage="Generate HTML report of the result.")]
    [switch]$HtmlReport,
    [Parameter(Mandatory=$false, HelpMessage="Number of threads to use for execution.")]
    [int32]$Threads = $env:NUMBER_OF_PROCESSORS
)

<#
TODO: 
        - $SiteServer should be validated - omg stupid hard
        - Exclude folders parameter in get-childitem?
        - publish to technet/github/psgallery
        - delete log entries for $result??
        - if given F:\ or \\server\f$ currently Get-AllPaths does not determine shared folders that match the path used
        - improve write-progress for -Process of main execution
        - test -altfoldersearch with use of where() and not where-object
#>

<#
    Define PSDefaultParameterValues and other variables
#>

$JobId = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'

# Write-CMLogEntry
$PSDefaultParameterValues["Write-CMLogEntry:Bias"]=(Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias)
$PSDefaultParameterValues["Write-CMLogEntry:Folder"]=($PSCommandPath | Split-Path -Parent)
$PSDefaultParameterValues["Write-CMLogEntry:FileName"]=(($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + ".log")
$PSDefaultParameterValues["Write-CMLogEntry:Enable"]=$Log.IsPresent
$PSDefaultParameterValues["Write-CMLogEntry:MaxLogFileSize"]=2MB
$PSDefaultParameterValues["Write-CMLogEntry:MaxNumOfRotatedLogs"]=0
$PSDefaultParameterValues["New-HTMLContent:SelectorColor"]="DeepSkyBlue"
$PSDefaultParameterValues["New-HTMLContent:HeaderBackGroundColor"]="DeepSkyBlue"
$PSDefaultParameterValues["New-HTMLTable:DisableColumnReorder"]=$true
$PSDefaultParameterValues["New-HTMLTable:ScrollX"]=$true
$PSDefaultParameterValues["New-HTMLTable:TextWhenNoData"]="None"

<#
    Define functions
#>

Function Write-CMLogEntry {
    <#
    .SYNOPSIS
    Write to log file in CMTrace friendly format.
    .DESCRIPTION
    Half of the code in this function is Cody Mathis's. I added log rotation and some other bits, with help of Chris Dent for some sorting and regex. Should find this code on the WinAdmins GitHub repo for configmgr.
    .OUTPUTS
    Writes to $Folder\$FileName and/or standard output.
    .LINK
    https://github.com/winadminsdotorg/SystemCenterConfigMgr
    #>
    param (
        [parameter(Mandatory = $true, HelpMessage = 'Value added to the log file.')]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $false, HelpMessage = 'Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('1', '2', '3')]
        [string]$Severity = 1,
        [parameter(Mandatory = $false, HelpMessage = "Stage that the log entry is occuring in, log refers to as 'component'.")]
        [ValidateNotNullOrEmpty()]
        [string]$Component,
        [parameter(Mandatory = $true, HelpMessage = 'Name of the log file that the entry will written to.')]
        [ValidateNotNullOrEmpty()]
        [string]$FileName,
        [parameter(Mandatory = $true, HelpMessage = 'Path to the folder where the log will be stored.')]
        [ValidateNotNullOrEmpty()]
        [string]$Folder,
        [parameter(Mandatory = $false, HelpMessage = 'Set timezone Bias to ensure timestamps are accurate.')]
        [ValidateNotNullOrEmpty()]
        [int32]$Bias,
        [parameter(Mandatory = $false, HelpMessage = 'Maximum size of log file before it rolls over. Set to 0 to disable log rotation.')]
        [ValidateNotNullOrEmpty()]
        [int32]$MaxLogFileSize = 0,
        [parameter(Mandatory = $false, HelpMessage = 'Maximum number of rotated log files to keep. Set to 0 for unlimited rotated log files.')]
        [ValidateNotNullOrEmpty()]
        [int32]$MaxNumOfRotatedLogs = 0,
        [parameter(Mandatory = $true, HelpMessage = 'A switch that enables the use of this function.')]
        [ValidateNotNullOrEmpty()]
        [switch]$Enable,
        [switch]$WriteHost
    )

    # Runs this regardless of $Enable value
    If ($WriteHost.IsPresent -eq $true) {
        Write-Host $Value
    }

    If ($Enable.IsPresent -eq $true) {
        # Determine log file location
        $LogFilePath = Join-Path -Path $Folder -ChildPath $FileName

        If ((([System.IO.FileInfo]$LogFilePath).Exists) -And ($MaxLogFileSize -ne 0)) {

            # Get log size in bytes
            $LogFileSize = [System.IO.FileInfo]$LogFilePath | Select-Object -ExpandProperty Length

            If ($LogFileSize -ge $MaxLogFileSize) {

                # Get log file name without extension
                $LogFileNameWithoutExt = $FileName -replace ([System.IO.Path]::GetExtension($FileName))

                # Get already rolled over logs
                $RolledLogs = "{0}_*" -f $LogFileNameWithoutExt
                $AllLogs = Get-ChildItem -Path $Folder -Name $RolledLogs -File

                # Sort them numerically (so the oldest is first in the list)
                $AllLogs = $AllLogs | Sort-Object -Descending { $_ -replace '_\d+\.lo_$' }, { [Int]($_ -replace '^.+\d_|\.lo_$') }
            
                ForEach ($Log in $AllLogs) {
                    # Get log number
                    $LogFileNumber = [int32][Regex]::Matches($Log, "_([0-9]+)\.lo_$").Groups[1].Value
                    switch (($LogFileNumber -eq $MaxNumOfRotatedLogs) -And ($MaxNumOfRotatedLogs -ne 0)) {
                        $true {
                            # Delete log if it breaches $MaxNumOfRotatedLogs parameter value
                            $DeleteLog = Join-Path $Folder -ChildPath $Log
                            [System.IO.File]::Delete($DeleteLog)
                        }
                        $false {
                            # Rename log to +1
                            $Source = Join-Path -Path $Folder -ChildPath $Log
                            $NewFileName = $Log -replace "_([0-9]+)\.lo_$",("_{0}.lo_" -f ($LogFileNumber+1))
                            $Destination = Join-Path -Path $Folder -ChildPath $NewFileName
                            [System.IO.File]::Copy($Source, $Destination, $true)
                        }
                    }
                }

                # Copy main log to _1.lo_
                $NewFileName = "{0}_1.lo_" -f $LogFileNameWithoutExt
                $Destination = Join-Path -Path $Folder -ChildPath $NewFileName
                [System.IO.File]::Copy($LogFilePath, $Destination, $true)

                # Blank the main log
                $StreamWriter = [System.IO.StreamWriter]::new($LogFilePath, $false)
                $StreamWriter.Close()
            }
        }

        # Construct time stamp for log entry
        switch -regex ($Bias) {
            '-' {
                $Time = [string]::Concat($(Get-Date -Format 'HH:mm:ss.fff'), $Bias)
            }
            Default {
                $Time = [string]::Concat($(Get-Date -Format 'HH:mm:ss.fff'), '+', $Bias)
            }
        }

        # Construct date for log entry
        $Date = (Get-Date -Format 'MM-dd-yyyy')
    
        # Construct context for log entry
        $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    
        # Construct final log entry
        $LogText = [string]::Format('<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">', $Value, $Time, $Date, $Component, $Context, $Severity, $PID)
    
        # Add value to log file
        try {
            $StreamWriter = [System.IO.StreamWriter]::new($LogFilePath, 'Append')
            $StreamWriter.WriteLine($LogText)
            $StreamWriter.Close()
        }
        catch [System.Exception] {
            Write-Warning -Message ("Unable to append log entry to {0} file. Error message: {1}" -f $FileName, $_.Exception.Message)
        }
    }
}

Function Get-CMContent {
    [CmdletBinding()]
    <#
    .SYNOPSIS
    Get all ConfigMgr objects that can hold content, i.e. content objects.
    .DESCRIPTION
    Using the ConfigMgr PoSH cmdlets, in the $Commands array, get all content objects and filter them to the given site code. 
    For each content object, create a PSCustomObject with the needed properties.
    Called by main body.
    .OUTPUTS
    System.Object.PSCustomObject
    #>
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string[]]$Commands,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false)]
        [string]$SiteServer,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false)]
        [string]$SiteCode
    )
    Begin {
        [hashtable]$ShareCache = @{}
    }
    Process {
        ForEach ($Command in $Commands) {

            Write-CMLogEntry -Value ("Gathering: {0}" -f $Command -replace "Get-CM") -Severity 1 -Component "GatherContentObjects"

            # Filter by site code
            $Command = $Command + " | Where-Object SourceSite -eq `"{0}`"" -f $SiteCode

            ForEach ($item in (Invoke-Expression $Command)) {
                switch -regex ($Command) {
                    "^Get-CMApplication.+" {
                        $AppMgmt = [xml]$item.SDMPackageXML | Select-Object -ExpandProperty AppMgmtDigest
                        ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
                            $SourcePaths = $DeploymentType.Installer.Contents.Content.Location
                            # Using ForEach-Object because even if $SourcePaths is null, it will iterate null once which is ideal here where deployment types can have no source path.
                            # Also, A deployment type can have more than 1 source path: for install and uninstall paths
                            $SourcePaths | ForEach-Object {
                                $SourcePath = $_

                                # Get every possible path
                                $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SiteServer $SiteServer

                                # Create content object PSObject with needed properties and add to array
                                $obj = [PSCustomObject]@{
                                    ContentType     = "Application"
                                    UniqueID        = $DeploymentType | Select-Object -ExpandProperty LogicalName
                                    Name            = "{0}::{1}" -f $item.LocalizedDisplayName,$DeploymentType.Title.InnerText
                                    SourcePath      = $SourcePath
                                    SourcePathFlag  = [int](Test-FileSystemAccess -Path $SourcePath -Rights Read)
                                    AllPaths        = $GetAllPathsResult[1]
                                }
                                $obj
                            }

                            # Maintaining cache of shared folders for servers encountered so far
                            $ShareCache = $GetAllPathsResult[0]

                            Write-CMLogEntry -Value ("{0} - {1} - {2} - {3} - {4} - {5}" -f $obj.ContentType,$obj.UniqueID,$obj.Name,$obj.SourcePath,$obj.SourcePathFlag,($obj.AllPaths.Keys -join ",")) -Severity 1 -Component "GatherContentObjects"
                        }
                    }
                    "^Get-CMDriver\s.+" { 
                        $SourcePath = $item.ContentSourcePath
                        # Get every possible path
                        $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SiteServer $SiteServer 
                        
                        # Create content object PSObject with needed properties and add to array
                        $obj = [PSCustomObject]@{
                            ContentType     = "Driver"
                            UniqueID        = $item.CI_ID
                            Name            = $item.LocalizedDisplayName
                            SourcePath      = $SourcePath
                            SourcePathFlag  = [int](Test-FileSystemAccess -Path $SourcePath -Rights Read)
                            AllPaths        = $GetAllPathsResult[1]
                        }
                        $obj

                        # Maintaining cache of shared folders for servers encountered so far
                        $ShareCache = $GetAllPathsResult[0]

                        Write-CMLogEntry -Value ("{0} - {1} - {2} - {3} - {4} - {5}" -f $obj.ContentType,$obj.UniqueID,$obj.Name,$obj.SourcePath,$obj.SourcePathFlag,($obj.AllPaths.Keys -join ",")) -Severity 1 -Component "GatherContentObjects"
                    }
                    default {
                        # OS images and boot iamges are absolute paths to files
                        If (($Command -match ("^Get-CMOperatingSystemImage.+")) -Or ($Command -match ("^Get-CMBootImage.+"))) {
                            $SourcePath = Split-Path $item.PkgSourcePath
                        }
                        Else {
                            $SourcePath = $item.PkgSourcePath
                        }
                        
                        $ContentType = ([Regex]::Matches($Command, "Get-CM([^\s]+)")).Groups[1].Value
                        # Get every possible path
                        $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SiteServer $SiteServer

                        # Create content object PSObject with needed properties and add to array
                        $obj = [PSCustomObject]@{
                            ContentType     = $ContentType
                            UniqueID        = $item.PackageId
                            Name            = $item.Name
                            SourcePath      = $SourcePath
                            SourcePathFlag  = [int](Test-FileSystemAccess -Path $SourcePath -Rights Read)
                            AllPaths        = $GetAllPathsResult[1]
                        }
                        $obj

                        # Maintaining cache of shared folders for servers encountered so far
                        $ShareCache = $GetAllPathsResult[0]

                        Write-CMLogEntry -Value ("{0} - {1} - {2} - {3} - {4} - {5}" -f $obj.ContentType,$obj.UniqueID,$obj.Name,$obj.SourcePath,$obj.SourcePathFlag,($obj.AllPaths.Keys -join ",")) -Severity 1 -Component "GatherContentObjects"
                    }   
                }
            }
            Write-CMLogEntry -Value ("Done gathering: {0}" -f ($Command -replace "Get-CM").Split(" ")[0]) -Severity 1 -Component "GatherContentObjects"
        }
    }
}

Function Get-AllPaths {
    <#
    .SYNOPSIS
    Determine all possible paths for a given path.
    .DESCRIPTION
    For a given path, determine all other possible path combinations that ultimately point back to $Path. 
    Useful to determine the local path for a given UNC path (in turn get the UNC path that uses the drive $ share), or if a there are multiple shared folders pointing to the same location. 
    Called by Get-CMContent.
    .OUTPUTS
    System.Object.List with always only two elements; $AllPaths (hashtable, the calculated list of "all paths" for the given $Path), and $Cache (hashtable, the shared folders cache).
    The first element ($Cache, hashtable) of the $result collection is dedicated to being a cache which will contain a list of all servers (key) and a hashtable (value) for a list of shared folder names and their local paths.
    The second element ($AllPath, hashtable) of the $result collection contains the list of all possible paths associated with the given $Path (key) and the NetBIOS server name (value) of which it belongs to. 
    .EXAMPLE
    PS C:\> Get-AllPaths -Path "\\SCCM\Applications$\7-zip" -Cache $SharedFolderCache -SiteServer "SCCM"

        Name                           Value                                                         
        ----                           -----                                                         
        192.168.175.11                 {Folder, UpdateServicesPackages, EasySetupPayload, SMSSIG$...}
        sccm                           {Folder, UpdateServicesPackages, EasySetupPayload, SMSSIG$...}
        sccm.acc.local                 {Folder, UpdateServicesPackages, EasySetupPayload, SMSSIG$...}
        \\192.168.175.11\Applications$ sccm
        \\192.168.175.11\F$\Applica... sccm
        \\sccm.acc.local\Applications$ sccm
        \\sccm\Applications$           sccm
        \\sccm\DiffFolder1$            sccm
        \\sccm.acc.local\F$\Applica... sccm
        \\192.168.175.11\DiffFolder1$  sccm
        F:\Applications                sccm
        \\sccm.acc.local\DiffFolder1$  sccm
        \\sccm\F$\Applications         sccm
    #>
    param (
        [string]$Path,
        [hashtable]$Cache,
        [string]$SiteServer
    )

    [System.Collections.Generic.List[Object]]$result = @()
    [hashtable]$AllPaths = @{}

    If (([string]::IsNullOrEmpty($Path) -eq $false) -And ($Path -notmatch "^[a-zA-Z]:\\$")) {
        $Path = $Path.TrimEnd("\")
    }

    ##### Determine path type

    switch ($true) {
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z]\$)$") {
            # Path that is \\server\f$
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$null
            $PathType = 4
            break
        }
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z]\$)(\\[a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)") {
            # Path that is \\server\f$\folder
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$Matches[3]
            $PathType = 3
            break
        }
        
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)$") {
            # Path that is \\server\share
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$null
            $PathType = 2
            break
        }
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)(\\[a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)") {
            # Path that is \\server\share\folder
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$Matches[3]
            $PathType = 1
            break
        }
        ($Path -match "^[a-zA-Z]:\\") {
            # Path that is local
            # Script does not determine UNC / shared folder paths if the content object source path is a local path
            $AllPaths.Add($Path, $SiteServer)
            $result.Add($Cache)
            $result.Add($AllPaths)
            return $result
        }
        ([string]::IsNullOrEmpty($Path) -eq $true) {
            # If there is no source path, just return now with $AllPaths empty
            $result.Add($Cache)
            $result.Add($AllPaths)
            return $result
        }
        default { 
            # Please share $Path with me if this is caught!
            # As a fail safe, abort
            $Message = "Unable to interpret path `"{0}`"" -f $Path
            Write-Warning $Message
            Write-CMLogEntry -Value $Message -Severity 2 -Component "GatherContentObjects"
            $AllPaths.Add($Path, $null)
            $result.Add($Cache)
            $result.Add($AllPaths)
            return $result
        }
    }

    ##### Determine FQDN, IP and NetBIOS

    # Only determine if you have a record
    # Might be annoying if $Server is an IP, unreachable and revese lookup succeeds
    If (Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue) {
        If ($Server -as [IPAddress]) {
            try {
                # Reverse lookup
                $FQDN = [System.Net.Dns]::GetHostEntry($Server) | Select-Object -ExpandProperty HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                # In case no record
                $FQDN = $null
            }
            $IP = $Server
        }
        Else {
            try {
                # Get FQDN even if $Server is FQDN, so we cut out $NetBIOS and resolve for $IP
                $FQDN = [System.Net.Dns]::GetHostByName($Server) | Select-Object -ExpandProperty HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                # In case no record
                $FQDN = $null
            }
            $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
        }
    }
    Else {
        # Won't be able to query Win32_Class if unreachable so no point continuing
        Write-CMLogEntry -Value ("Server `"{0}`" is unreachable" -f $Server) -Severity 2 -Component "GatherContentObjects"
        $AllPaths.Add($Path, $null)
        $result.Add($Cache)
        $result.Add($AllPaths)
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
            # Add null so on the next encounter of a server from a given UNC path, we don't wastefully try again
            $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object {
                $Cache.Add($_, $null)
            }
            Write-CMLogEntry -Value ("Could not update cache because could not get shared folders from: `"{0}`"" -f $FQDN) -Severity 2 -Component "GatherContentObjects"
        }
    }

    ##### Build the AllPaths property

    [System.Collections.Generic.List[String]]$AllPathsArr = @()

    $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object -Process {
        $AltServer = $_
        If ($Cache.$AltServer -ne $null) {
            # Get the share's local path
            $LocalPath = $Cache.$AltServer.GetEnumerator().Where( { $_.Key -eq $ShareName } ) | Select-Object -ExpandProperty Value
        }
        Else {
            $LocalPath = $null
        }
        # If \\server\f$ then $LocalPath is "F:"
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            If ($PathType -match "1|2") {
                # Add \\server\f$\path\to\shared\folder\on\disk
                $AllPathsArr.Add(("\\$($AltServer)\$($LocalPath)$($ShareRemainder)" -replace ':', '$'))
                # Get other shared folders that point to the same path and add them to the AllPaths array
                $SharesWithSamePath = $Cache.$AltServer.GetEnumerator().Where( { $_.Value -eq $LocalPath } ) | Select-Object -ExpandProperty Key
                ForEach ($AltShareName in $SharesWithSamePath) {
                    $AllPathsArr.Add("\\$($AltServer)\$($AltShareName)$($ShareRemainder)")
                }
            }  
        }
        Else {
            Write-CMLogEntry -Value ("Could not resolve share `"{0}`" on `"{1}`" from cache, either because it does not exist or could not query Win32_Share on server" -f $ShareName,$_) -Severity 2 -Component "GatherContentObjects"
        }
        # Add the original path again but with the alternate server (FQDN / NetBIOS / IP)
        $AllPathsArr.Add("\\$($AltServer)\$($ShareName)$($ShareRemainder)")
    } -End {
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            # Either of the below are important in case user is running local to site server and gave local path as $SourcesLocation
            If (($LocalPath -match "^[a-zA-Z]:$") -And ($PathType -match "2|4")) {
                # Match if just a drive letter (WHY?!) and add it to AllPaths array
                # This occurs if path type is 2 and the share points to root of a volume
                $AllPathsArr.Add("$($LocalPath)\")
            }
            Else {
                # Add the local path to AllPaths array
                $AllPathsArr.Add("$($LocalPath)$($ShareRemainder)")
            }
        }
    }

    # Add all that's inside the AllPaths array to the AllPaths hashtable
    # Unfotunately adding stuff to hash table that already exists in there can be noisy to stderr in console
    ForEach ($item in $AllPathsArr) {
        If (($AllPaths.Keys -notcontains $item) -eq $true) {
            $AllPaths.Add($item, $NetBIOS)
        }
    }

    $result.Add($Cache)
    $result.Add($AllPaths)
    return $result
}

Function Get-AllSharedFolders {
    <#
    .SYNOPSIS
    Get all shared folders hosted on a server.
    .DESCRIPTION
    Query Win32_Share WMI class on $Server and return a hashtable result. 
    Called by Get-AllPaths.
    .OUTPUTS
    System.Object.Hashtable where a list of shared folder names (key) and their local paths (value).
    #>
    Param([String]$Server)

    [hashtable]$AllShares = @{}

    try {
        $Shares = (Get-WmiObject -ComputerName $Server -Class Win32_Share -ErrorAction Stop).Where( {-not [string]::IsNullOrEmpty($_.Path)} )
        ForEach ($Share in $Shares) {
            # The TrimEnd method is only really concerned for drive letter shares
            # as they're usually stored as f$ = "F:\" and this messes up Get-AllPaths a little
            $AllShares.Add($Share.Name, $Share.Path.TrimEnd("\"))
        }
    }
    catch {
        $AllShares = $null
    }

    return $AllShares
}

Function Get-AllFolders {
    <#
    .SYNOPSIS
    Recrusively get all folders under $Path.
    .DESCRIPTION
    Get all folders in $Path. By default this function escapes the max path limit by prefixing $Path with the following: "\\?\UNC". This is what _mostly_ the driver for the PoSH 5.1 requirement.
    Called by main body.
    .OUTPUTS
    System.Object.Generic.List[String] of folder full names.
    #>
    Param(
        [string]$Path,
        [bool]$DiffFolderSearch
    )

    # Prefix the path the user gives us with \\?\ to avoid the 260 MAX_PATH limit
    # More info https://docs.microsoft.com/en-us/windows/desktop/fileio/naming-a-file#maximum-path-length-limitation
    switch ($true) { 
        ($Path -match "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z0-9\\`~!@#$%^&(){}\'._ -]+") {
            # Matches if it's a UNC path
            # Could have queried .IsUnc property on [System.Uri] object but I wanted to verify user hadn't first given us \\?\ path type
            $Path = $Path -replace "^\\\\", "\\?\UNC\"
            break
        }
        ($Path -match "^[a-zA-Z]:\\") {
            # Matches if starts with drive letter
            $Path = "\\?\" + $Path
            break
        }
        default {
            $Message = "Couldn't determine path type for `"{0}`" so might have problems accessing folders that breach MAX_PATH limit, quitting..." -f $Path
            Write-CMLogEntry -Value $Message -Severity 2 -Component "GatherFolders"
            throw $Message
        }
    }
    
    # Recursively get all folders
    If ($DiffFolderSearch -eq $true) {
        [System.Collections.Generic.List[String]]$Folders = Start-AltFolderSearch -FolderName $Path
    }
    Else {
        try {
            [System.Collections.Generic.List[String]]$Folders = Get-ChildItem -LiteralPath $Path -Directory -Recurse | Select-Object -ExpandProperty FullName
        }
        catch {
            $Message = "Consider using -AltFolderSearch, quiting..."
            Write-CMLogEntry -Value $Message -Severity 3 -Component "GatherFolders"
            throw $Message
        }
    }

    If ([string]::IsNullOrEmpty($Folders) -eq $true) {
        [System.Collections.Generic.List[String]]$Folders = @($Path)
    }
    Else {
        $Folders.Add($Path)
    }

    # Undo the \\?\ prefix
    switch ($true) {
        ($Path -match "^\\\\\?\\UNC\\") {
            # Matches if starts with \\?\UNC\
            $Folders = $Folders -replace [Regex]::Escape("\\?\UNC\"), "\\"
            break
        }
        ($Path -match "^\\\\\?\\[a-zA-Z]{1}:\\") {
            # Matches if starts with \\?\A:\ (A is just an example drive letter used)
            $Folders = $Folders -replace [Regex]::Escape("\\?\"), ""
            break
        }
        default {
            # For some reason, couldn't undo \\?\ prefix. If you get this, please share $Path with me!
            # No big deal though, can keep going. $SourcesLocation will stay as what the user gave in the parent scope
            Write-CMLogEntry -Value ("Couldn't reset {0}" -f $Path) -Severity 3 -Component "GatherFolders"
        }
    }
    
    $Folders = $Folders | Sort-Object

    return $Folders
}
Function Start-AltFolderSearch {
    <#
    .SYNOPSIS
    Get all folders under $FolderName, but not recursively.
    .DESCRIPTION
    Get all folders under $FolderName but does not recursively get all folders for each and every child funder. 
    This exists because in some environments Get-ChildItem would throw an exception "Not enough quota is available to process this command.". FullyQualifiedErrorId: "DirIOError,Microsoft.PowerShell.Commands.GetChildItemCommand".
    While investigating the exception was thrown at around the 50k size of any collection type and packet traces showed SMBv1 packets returning similar exception message as by PoSH, some sort of quota limit.
    Further testing on different storage systems using SMBv1 this exception was not reproducable.
    Massive thanks to Chris Kibble for coming up with this work around and time to help troubleshoot!
    Called by Get-AllFolders.
    .OUTPUTS
    System.Object.Generic.List[String] of folder full names.
    #>
    Param([string]$FolderName)

    # Annoyingly, Get-ChildItem with forced output to an arry @(Get-ChildItem ...) can return an explicit
    # $null value for folders with no subfolders, causing the for loop to indefinitely iterate through
    # working dir when it reaches a null value, so is null check is needed
    [System.Collections.Generic.List[String]]$Folders = @((Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName).Where( { [string]::IsNullOrEmpty($_) -eq $false } ))

    ForEach($Folder in $Folders) {
        $Folders.Add($(Start-AltFolderSearch -FolderName $Folder))
    }

    return $Folders
}

Function Test-FileSystemAccess {
    <#
    .SYNOPSIS
        Check for file system access on a given folder.
    .OUTPUTS
        [System.Enum]
        ERROR_SUCCESS (0)
        ERROR_PATH_NOT_FOUND (3)
        ERROR_ACCESS_DENIED (5)
        ERROR_ELEVATION_REQUIRED (740)
    .NOTES
        Authors:    Patrick Seymour / Adam Cook
        Contact:    @pseymour / @codaamok
    #>
    param
    (
        [Parameter(Mandatory=$false)]
        [string]$Path,
        [Parameter(Mandatory=$true)]
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    enum FileSystemAccessState {
        ERROR_SUCCESS
        ERROR_PATH_NOT_FOUND = 3
        ERROR_ACCESS_DENIED = 5
        ERROR_ELEVATION_REQUIRED = 740
    }

    [System.Security.Principal.WindowsIdentity]$currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    [System.Security.Principal.WindowsPrincipal]$currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $IsElevated = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $IsInAdministratorsGroup = $currentIdentity.Claims.Value -contains "S-1-5-32-544"

    if ([System.IO.Directory]::Exists($Path))
    {
        try
        {
            [System.Security.AccessControl.FileSystemSecurity]$security = (Get-Item -Path ("FileSystem::{0}" -f $Path) -Force).GetAccessControl()
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
                            return [FileSystemAccessState]::ERROR_SUCCESS
                        }
                    }
                }

                if (($IsElevated -eq $false) -And ($IsInAdministratorsGroup -eq $true) -And ($rules.Where( { ($_.IdentityReference -eq "S-1-5-32-544") -And ($_.FileSystemRights.HasFlag($Rights)) } )))
                {
                    # At this point we were able to read ACL and verify Administrators group access, likely because we were qualified by the object set as owner
                    return [FileSystemAccessState]::ERROR_ELEVATION_REQUIRED
                }
                else
                {
                    return [FileSystemAccessState]::ERROR_ACCESS_DENIED
                }

            }
            else
            {
                return [FileSystemAccessState]::ERROR_ACCESS_DENIED
            }
        }
        catch
        {
            return [FileSystemAccessState]::ERROR_ACCESS_DENIED
        }
    }
    else
    {
        return [FileSystemAccessState]::ERROR_PATH_NOT_FOUND
    }
}

function Measure-ChildItem {
    <#
    .SYNOPSIS
    Recursively measures the size of a directory.
    .NOTES
    Author: Chris Dent (indented-automation) https://github.com/indented-automation
    Source: https://github.com/steviecoaster/PSSysadminToolkit/blob/Dev/Public/Measure-ChildItem.ps1
    MIT license. http://www.opensource.org/licenses/MIT
    #>

    [CmdletBinding()]
    param (
        # The path to measure the size of. Accepts pipeline input. By default the size of the current working directory is measured.
        [Parameter(Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [String]$Path = $pwd,

        # The units sizes should be displayed in. By default, sizes are displayed in Bytes.
        [ValidateSet('B', 'KB', 'MB', 'GB', 'TB')]
        [String]$Unit = 'B',

        # When rounding, the number of digits to display after a decimal point. By defaut sizes are rounded to two decimal places.
        [ValidateRange(0, 28)]
        [Int32]$Digits = 2,

        # Return the size value only, discards file, and directory counts and path information.
        [Switch]$ValueOnly
    )

    begin {
        if (-not ('SC.IO.FileSearcher' -as [Type])) {
            Add-Type '
                using System;
                using System.Collections.Generic;
                using System.IO;
                using System.Runtime.InteropServices;

                namespace SC.IO
                {
                    [StructLayout(LayoutKind.Sequential)]
                    public struct FILETIME
                    {
                        public uint dwLowDateTime;
                        public uint dwHighDateTime;
                    };

                    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
                    public struct WIN32_FIND_DATA
                    {
                        public FileAttributes dwFileAttributes;
                        public FILETIME ftCreationTime;
                        public FILETIME ftLastAccessTime;
                        public FILETIME ftLastWriteTime;
                        public int nFileSizeHigh;
                        public int nFileSizeLow;
                        public int dwReserved0;
                        public int dwReserved1;
                        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                        public string cFileName;
                        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
                        public string cAlternate;
                    }

                    public class UnsafeNativeMethods
                    {
                        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
                        public static extern IntPtr FindFirstFile(string lpFileName, out WIN32_FIND_DATA lpFindFileData);

                        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
                        public static extern IntPtr FindFirstFileExW(
                            string              lpFileName,
                            int                 fInfoLevelId,
                            out WIN32_FIND_DATA lpFindFileData,
                            int                 fSearchOp,
                            IntPtr              lpSearchFilter,
                            int                 dwAdditionalFlags
                        );

                        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
                        public static extern bool FindNextFile(IntPtr hFindFile, out WIN32_FIND_DATA lpFindFileData);

                        [DllImport("kernel32.dll", SetLastError = true)]
                        [return: MarshalAs(UnmanagedType.Bool)]
                        public static extern bool FindClose(IntPtr hFindFile);
                    }

                    public class FileSearcher
                    {
                        private static uint convertToUInt(int value)
                        {
                            return BitConverter.ToUInt32(
                                BitConverter.GetBytes(value),
                                0
                            );
                        }

                        private static long convertToLong(int value)
                        {
                            return (long)(convertToUInt(value) << 32);
                        }

                        public static long[] MeasureItem(string path, bool recurse, long[] itemData)
                        {
                            if (itemData == null)
                            {
                                itemData = new long[]{ 0, 0, 0 };
                            }

                            string searchPath;
                            if (path.StartsWith(@"\\"))
                            {
                                searchPath = String.Format(@"\\?\UNC\{0}\*", path.Substring(2));
                            }
                            else
                            {
                                searchPath = String.Format(@"\\?\{0}\*", path);
                            }

                            WIN32_FIND_DATA findData = new WIN32_FIND_DATA();
                            IntPtr findHandle = UnsafeNativeMethods.FindFirstFileExW(searchPath, 1, out findData, 0, IntPtr.Zero, 0);
                            do
                            {
                                if (findData.dwFileAttributes.HasFlag(FileAttributes.Directory))
                                {
                                    if (recurse && findData.cFileName != "." && findData.cFileName != "..")
                                    {
                                        itemData[2]++;
                                        itemData = MeasureItem(
                                            Path.Combine(path, findData.cFileName),
                                            recurse,
                                            itemData
                                        );
                                    }
                                }
                                else
                                {
                                    itemData[0] += convertToLong(findData.nFileSizeHigh) + (long)convertToUInt(findData.nFileSizeLow);
                                    itemData[1]++;
                                }
                            } while (UnsafeNativeMethods.FindNextFile(findHandle, out findData));
                            UnsafeNativeMethods.FindClose(findHandle);

                            return itemData;
                        }
                    }
                }
            '
        }

        $power = ('B', 'KB', 'MB', 'GB', 'TB').IndexOf($Unit.ToUpper())
        $denominator = [Math]::Pow(1024, $power)
    }

    process {
        $Path = $pscmdlet.GetUnresolvedProviderPathFromPSPath($Path).TrimEnd('\')

        $itemData = [SC.IO.FileSearcher]::MeasureItem($Path, $true, $null)

        if ($ValueOnly) {
            [Math]::Round(($itemData[0] / $denominator), $Digits)
        } else {
            [PSCustomObject]@{
                Path           = $Path
                Size           = [Math]::Round(($itemData[0] / $denominator), $Digits)
                FileCount      = $itemData[1]
                DirectoryCount = $itemData[2]
            }
        }
    }
}

Function Set-CMDrive {
    <#
    .SYNOPSIS
    Import ConfigMgr module, create ConfigrMgr PS drive and set location to it.
    .DESCRIPTION
    Set current working directory to site code for access to ConfigMgr cmdlets. Some validation is in place to verify the site code marrys up to be of $Server.
    Called by main body.
    #>
    Param(
        [string]$SiteCode,
        [string]$Server,
        [string]$Path
    )

    # Import the ConfigurationManager.psd1 module 
    If((Get-Module ConfigurationManager) -eq $null) {
        try {
            Import-Module ("{0}\..\ConfigurationManager.psd1" -f $ENV:SMS_ADMIN_UI_PATH)
        }
        catch {
            $Message = "Failed to import Configuration Manager module"
            Write-CMLogEntry -Value $Message -Severity 3 -Component "Initialisation"
            throw $Message
        }
    }

    try {
        # Connect to the site's drive if it is not already present
        If((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Server -ErrorAction Stop | Out-Null
        }
        # Set the current location to be the site code.
        Set-Location ("{0}:\" -f $SiteCode) -ErrorAction Stop

        # Verify given sitecode
        If((Get-CMSite -SiteCode $SiteCode | Select-Object -ExpandProperty SiteCode) -ne $SiteCode) { throw }

    } 
    catch {
        If((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -ne $null) {
            Set-Location $Path
            Remove-PSDrive -Name $SiteCode -Force
        }
        $Message = "Failed to create New-PSDrive with site code `"{0}`" and server `"{1}`"" -f $SiteCode, $Server
        Write-CMLogEntry -Value $Message -Severity 3 -Component "Initialisation"
        throw $Message
    }

}

Write-CMLogEntry -Value "Starting" -Severity 1 -Component "Initilisation" -WriteHost

# To calculate total runtime
$StartTime = Get-Date

# Write all parameters passed to script to log
ForEach($item in $PSBoundParameters.GetEnumerator()) {
    Write-CMLogEntry -Value ("- {0}: {1}" -f $item.Key, $item.Value) -Severity 1 -Component "Initilisation"
}

# If user has given local path for $SourcesLocation, need to ensure we don't produce false positives where a similar folder structure exists on the remote machine and site server. e.g. packages let you specify local path on site server
If ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($env:COMPUTERNAME -ne $SiteServer)) {
    $Message = "Won't be able to determine unused folders with given local path while running remotely from site server, quitting"
    Write-CMLogEntry -Value $Message -Severity 2 -Component "Initilisation" -WriteHost
    throw $Message
}

# Import PSWriteHtml module report if -HtmlReport is present
If ($HtmlReport.IsPresent -eq $true) {
    try {
        Import-Module PSWriteHTML -ErrorAction Stop
    }
    catch {
        $Message = "Unable to import PSWriteHtml module: {0}" -f $error[0].Exception.Message
        Write-CMLogEntry -Value $Message -Severity 3 -Component "Initialisation"
        throw $Message
    }
    [version]$moduleVersion = (Get-Module PSWriteHTML | Sort-Object Version -Descending | Select-Object -ExpandProperty Version)[0]
    If($moduleVersion -lt [version]"0.0.44") {
        $Message = "PSWriteHtml version is too old ({0}).  Requires 0.0.44+." -f $moduleVersion.ToString()
        Write-CMLogEntry -Value $Message -Severity 3 -Component "Initialisation"
        throw $Message
    }
}

# Build the $Commands array ready for Get-CMContent
switch ($true) {
    ($Packages.IsPresent -eq $true) {
        [array]$Commands += "Get-CMPackage"
    }
    ($Applications.IsPresent -eq $true) {
        [array]$Commands += "Get-CMApplication"
    }
    ($Drivers.IsPresent -eq $true) {
        [array]$Commands += "Get-CMDriver"
    }
    ($DriverPackages.IsPresent -eq $true) {
        [array]$Commands += "Get-CMDriverPackage"
    }
    ($OSImages.IsPresent -eq $true) {
        [array]$Commands += "Get-CMOperatingSystemImage"
    }
    ($OSUpgradeImages.IsPresent -eq $true) {
        [array]$Commands += "Get-CMOperatingSystemInstaller"
    }
    ($BootImages.IsPresent -eq $true) {
        [array]$Commands += "Get-CMBootImage"
    }
    ($DeploymentPackages.IsPresent -eq $true) {
        [array]$Commands += "Get-CMSoftwareUpdateDeploymentPackage"
    }
    default {
        [array]$Commands = "Get-CMPackage", "Get-CMApplication", "Get-CMDriver", "Get-CMDriverPackage", "Get-CMOperatingSystemImage", "Get-CMOperatingSystemInstaller", "Get-CMBootImage", "Get-CMSoftwareUpdateDeploymentPackage"
    }
}

# Get NetBIOS of given $SiteServer parameter so it's similar format as $env:COMPUTERNAME used in body during folder/content object for loop
# And also for value pair in each content objects .AllPaths property (hashtable)
If ($SiteServer -as [IPAddress]) {
    $FQDN = [System.Net.Dns]::GetHostEntry($SiteServer) | Select-Object -ExpandProperty HostName
}
Else {
    $FQDN = [System.Net.Dns]::GetHostByName($SiteServer) | Select-Object -ExpandProperty HostName
}
$SiteServer = $FQDN.Split(".")[0]

# Gather folders

Write-CMLogEntry -Value ("Gathering folders: {0}" -f $SourcesLocation) -Severity 1 -Component "GatherFolders" -WriteHost
If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 0 -Status ("Gathering all folders at: {0}" -f $SourcesLocation) }
$AllFolders = Get-AllFolders -Path $SourcesLocation -AltFolderSearch $AltFolderSearch.IsPresent
Write-CMLogEntry -Value ("Number of gathered folders: {0}" -f $AllFolders.count) -Severity 1 -Component "GatherFolders" -WriteHost

# Gather content objects

$OriginalPath = Get-Location | Select-Object -ExpandProperty Path
Set-CMDrive -SiteCode $SiteCode -Server $SiteServer -Path $OriginalPath

Write-CMLogEntry -Value ("Gathering content objects: {0}" -f ($Commands -replace "Get-CM" -join ", ")) -Severity 1 -Component "GatherContentObjects" -WriteHost
If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 33 -Status ("Gathering CM content objects: {0}" -f ($Commands -replace "Get-CM" -join ", ")) }
$AllContentObjects = Get-CMContent -Commands $Commands -SiteServer $SiteServer -SiteCode $SiteCode
Write-CMLogEntry -Value ("Number of gathered content objects: {0}" -f $AllContentObjects.count) -Severity 1 -Component "GatherContentObjects" -WriteHost

Set-Location $OriginalPath

$AllFolders | ForEach-Object -Begin {

    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 66 -Status "Determining unused folders" }
    Write-CMLogEntry -Value ("Determining unused folders, using {0} threads" -f $Threads) -Severity 1 -Component "Processing" -WriteHost
    
    # Make Test-FileSystemAccess function available to all runspaces
    $Definition = Get-Content Function:\Test-FileSystemAccess -ErrorAction Stop
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Test-FileSystemAccess', $Definition
    $initialSessionState = [InitialSessionState]::CreateDefault()
    $InitialSessionState.Commands.Add($SessionStateFunction)

    # Create runspace pool, initialise the results array and script block it'll churn through
    $RSPool = [RunspaceFactory]::CreateRunspacePool(1, $Threads, $InitialSessionState, $Host)
    $RSPool.ApartmentState = "MTA"
    $RSPool.Open()
    [System.Collections.Generic.List[Object]]$RSResults = @()
    $RSScriptBlock = {
        Param (
            [string]$RSFolder,
            [System.Collections.Generic.List[Object]]$RSAllContentObjects
        )

        # Initialise the essentials
        [System.Collections.Generic.List[String]]$UsedBy = @()
        $IntermediatePath = $false
        $NotUsed = $false

        switch ([int](Test-FileSystemAccess -Path $RSFolder -Rights Read)) {
            5 {
                $UsedBy.Add("Access denied")
            }
            740 {
                $UsedBy.Add("Access denied (elevation required)")
            }
        }

        # Still continue anyway, despite access denied, because we can still determine if it's an exact match or intermediate path of a content object

        # Filtered to exclude SourcePathFlag 3 so we can exclude false positives
        # e.g. if content objects uses \\server\share\folder1\folder2 and $SourcesLocation is \\server\share\folder1
        # but SourcePathFlag is 3, this could report \\server\share\folder1 as intermediate where it may not be
        # Plus, no point iterating over them if we already know that the SourcePath isn't resolvable
        ForEach ($ContentObject in ($RSAllContentObjects.Where( {$_.SourcePathFlag -ne 3} ))) {

            switch($true) {
                ([string]::IsNullOrEmpty($ContentObject.SourcePath) -eq $true) {
                    # Content object source path is empty, no point continuing
                    break
                }
                ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($ContentObject.AllPaths.$RSFolder -eq $env:COMPUTERNAME)) {
                    # Content object source path is on local file system to the site server
                    $UsedBy.Add($ContentObject.Name)
                    break
                }
                (($ContentObject.AllPaths.Keys -contains $RSFolder) -eq $true) {
                    # By default the ContainsKey method ignores case
                    # A match has been found within the AllPaths property of the content object
                    $UsedBy.Add($ContentObject.Name)
                    break
                }
                (($ContentObject.AllPaths.Keys -match [Regex]::Escape($RSFolder)).Count -gt 0) {
                    # If any of the content object paths start with $RSFolder
                    $IntermediatePath = $true
                    break
                }
                ($ContentObject.AllPaths.Keys.Where{$RSFolder.StartsWith($_, "CurrentCultureIgnoreCase")}.Count -gt 0) {
                    # If $RSFolder starts wtih any of the content object paths
                    $IntermediatePath = $true
                    break
                }
                default {
                    # Folder isn't known to any content objects
                    $NotUsed = $true
                }
            }

        }

        switch ($true) {
            ($UsedBy.count -gt 0) {
                $UsedBy = $UsedBy -join ", "
                break
            }
            ($IntermediatePath -eq $true) {
                $UsedBy = "An intermediate folder (sub or parent folder)"
                break
            }
            ($NotUsed -eq $true) {
                $UsedBy = "Not used"
                break
            }
        }

        [PSCustomObject]@{
            Folder  = $RSFolder
            UsedBy  = $UsedBy -join ", "
        }

    }

    If ($NoProgress.IsPresent -eq $false) {
        If ($AllFolders.count -gt 150) {
            [Int32]$FolderInterval = $AllFolders.count * 0.01
        }
        Else {
            $FolderInterval = 2
        }
    }

    Write-CMLogEntry -Value "Adding jobs to queue" -Severity 1 -Component "Processing" -WriteHost

} -Process {

    $Folder = $_

    If ($NoProgress.IsPresent -eq $false) {
        If (($AllFolders.IndexOf($Folder) % $FolderInterval) -eq 0) {
            [Int32]$Percentage = ($AllFolders.IndexOf($Folder) / $AllFolders.count * 100)
            Write-Progress -Id 2 -Activity "Adding jobs to queue" -PercentComplete $Percentage -Status ("{0}% complete" -f $Percentage) -ParentId 1
        }
    }

    $Runspace = [PowerShell]::Create()
    $null = $Runspace.AddScript($RSScriptBlock)
    $null = $Runspace.AddArgument($Folder)
    $null = $Runspace.AddArgument($AllContentObjects)
    $Runspace.Runspacepool = $RSPool
    $RSResults.Add( [PSCustomObject]@{ Pipe = $Runspace; Status = $Runspace.BeginInvoke() } )
    
} -End {
    
    Write-CMLogEntry -Value "Waiting for jobs to complete" -Severity 1 -Component "Processing" -WriteHost

    [System.Collections.Generic.List[Object]]$Result = @()

    # Process runspaces, wait for their results and clean up when complete
    $TotalRunspaces = $RSResults | Measure-Object | Select-Object -ExpandProperty Count
    while ($RSResults.Status -ne $null) {
        If ($NoProgress.IsPresent -eq $false) { 
            $TotalNotComplete = $RSResults.Where( { $_.Status -eq $null } ) | Measure-Object | Select-Object -ExpandProperty Count
            Write-Progress -Id 2 -Activity "Evaluating folders" -Status ("{0} folders remaining" -f ($TotalRunspaces-$TotalNotComplete)) -PercentComplete ($TotalNotComplete/$TotalRunspaces * 100) -ParentId 1
        }
        $Completed = $RSResults.Where( { $_.Status.IsCompleted -eq $true } )
        ForEach ($item in $Completed) {
            # Reference index 0 so we can grab the PSCustomobject inside the PSDataCollection object
            $Result.Add(($item.Pipe.EndInvoke($item.Status)[0]))
            $item.Pipe.Dispose()
            $item.Status = $null
        }
        Start-Sleep -Seconds 2
    }

    # Clean up runspace pool
    $RSPool.Dispose()

    Write-CMLogEntry -Value "Done determining unused folders" -Severity 1 -Component "Processing" -WriteHost

    # Update Write-Progress
    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Evaluating folders" -Completed -ParentId 1 }
    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 100 -Status "Finishing up" }

    Write-CMLogEntry -Value "Calculating used disk space by unused folders" -Severity 1 -Component "Exit" -WriteHost
    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Calculating used disk space by unused folders" -PercentComplete 0 -ParentId 1 }

    # Calculate total MB used for each "Not used" folder
    # This is wasteful if -HtmlReport is not specified, but I really wanted $SummaryNotUsedFolders total in end result stats
    $NotUsedFolders = $Result.Where( { $_.UsedBy -eq "Not used" } )

    # Calculate total MB used on size unused by ConfigMgr
    If ($NotUsedFolders.count -eq 0) {
        # PSCustomObject created so that when $NotUsedFolders is blank, PSWriteHtml won't print warnings because of missing properties when trying to create merge headers
        $SummaryNotUsedFolders = [PSCustomObject]@{
            Path            = 0
            Size            = 0
            FileCount       = 0
            DirectoryCount  = 0
        }
        $SummaryNotUsedFoldersMB,$SummaryNotUsedFoldersFileCount,$SummaryNotUsedFoldersDirectoryCount = 0,0,0
    }
    Else {
        $SummaryNotUsedFolders = $NotUsedFolders | Sort-Object Folder | ForEach-Object {
            $current = $_
            If (($previous) -And ($current.Folder.StartsWith($previous.Folder))) {
                # Do nothing
            }
            Else {
                $previous = $current
                $current.Folder | Measure-ChildItem -Unit MB -Digits 2
            }
        }
        Write-CMLogEntry -Value "Done calculating used disk space by unused folders" -Severity 1 -Component "Exit" -WriteHost
        $SummaryNotUsedFoldersMB = [Math]::Round(($SummaryNotUsedFolders | Measure-Object Size -Sum | Select-Object -ExpandProperty Sum), 2)
        $SummaryNotUsedFoldersFileCount = $SummaryNotUsedFolders | Measure-Object FileCount -Sum | Select-Object -ExpandProperty Sum
        $SummaryNotUsedFoldersDirectoryCount = $SummaryNotUsedFolders | Measure-Object DirectoryCount -Sum | Select-Object -ExpandProperty Sum
    }

    # Write $Result to log file
    # I know Write-CMLogEntry has Enabled parameter but having it here too just makes sense - to save the gazillion of loops for something that may be disabled anyway
    # May consider deleting this section, enough about the result is written to file
    If ($Log.IsPresent -eq $true) {
        $Message = "Writing result to log file"
        If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity $Message -PercentComplete 25 -ParentId 1 }
        Write-CMLogEntry -Value $Message -Severity 1 -Component "Processing" -WriteHost
        ForEach ($item in $Result) {
            switch -regex ($item.UsedBy) {
                "Access denied" {
                    $Severity = 2
                }
                default {
                    $Severity = 1
                }
            }
            Write-CMLogEntry -Value ($item.Folder + ": " + $item.UsedBy) -Severity $Severity -Component "Processing"
        }
    }

    # Export $Result to file
    If ($ExportReturnObject.IsPresent -eq $true) {
        try {
            Write-CMLogEntry -Value "Exporting PowerShell return object" -Severity 1 -Component "Exit" -WriteHost
            If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Exporting PowerShell return object" -PercentComplete 50 -ParentId 1 }
            Export-Clixml -LiteralPath (($PSCommandPath | Split-Path -Parent) + "\" + ($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + "_result.xml") -InputObject $Result
            Write-CMLogEntry -Value "Done exporting PowerShell return object" -Severity 1 -Component "Exit" -WriteHost
        }
        catch {
            Write-CMLogEntry -Value ("Failed to export PowerShell object: {0}" -f $error[0].Exception.Message) -Severity 3 -Component "Exit" -WriteHost
        }
    }

    # Export $AllContentObjects to file
    If ($ExportCMContentObjects.IsPresent -eq $true) {
        try {
            Write-CMLogEntry -Value "Exporting PowerShell ConfigMgr content objects object" -Severity 1 -Component "Exit" -WriteHost
            If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Exporting PowerShell ConfigMgr content objects object" -PercentComplete 75 -ParentId 1 }
            Export-Clixml -LiteralPath (($PSCommandPath | Split-Path -Parent) + "\" + ($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + "_cmobjects.xml") -InputObject $AllContentObjects
            Write-CMLogEntry -Value "Done exporting PowerShell ConfigMgr content objects object" -Severity 1 -Component "Exit" -WriteHost
        }
        catch {
            Write-CMLogEntry -Value ("Failed to export PowerShell object: {0}" -f $error[0].Exception.Message) -Severity 3 -Component "Exit" -WriteHost
        }
    }

    # Write $Result to HTML using PSWriteHTML
    If ($HtmlReport.IsPresent -eq $true) {
        try {
            Write-CMLogEntry -Value "Creating HTML report" -Severity 1 -Component "Exit" -WriteHost
            If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Creating HTML report" -PercentComplete 100 -ParentId 1 }

            New-HTML -TitleText ("Get-CMUnusedSources - {0}" -f $JobId) -UseCssLinks:$true -UseJavaScriptLinks:$true -FilePath (($PSCommandPath | Split-Path -Parent) + "\" + ($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + ".html") -ShowHTML {
                New-HTMLTabOptions -SlimTabs
                $Title = "All folders"
                New-HTMLTab -Name $Title {
                    New-HTMLContent -HeaderText $Title {
                        New-HTMLPanel {
                            New-HTMLTable -DataTable $Result {
                                New-HTMLTableCondition -Name "UsedBy" -Type string -Operator contains -Value "Access denied" -BackgroundColor Orange -Row
                            } -PreContent {
                                ("<span style='font-size: 1.2em; margin-left: 1em;'>All of the folders under `"{0}`" and their UsedBy status, same as what's returned by the script.</span>" -f $SourcesLocation)
                            }
                        }
                    }
                }
                $Title = "Summary of not used folders"
                New-HTMLTab -Name $Title {
                    New-HTMLContent -HeaderText $Title {
                        New-HTMLPanel {
                            New-HTMLTable -DataTable ($SummaryNotUsedFolders | Select-Object Path, @{Label="Size (MB)"; Expression={$_.Size}}, FileCount, DirectoryCount) {
                                New-HTMLTableHeader -Names "Size (MB)" -Title ("{0}MB" -f $SummaryNotUsedFoldersMB) -Color White -FontWeight Bold -Alignment left -BackGroundColor LimeGreen
                                New-HTMLTableHeader -Names "FileCount" -Title $SummaryNotUsedFoldersFileCount -Color White -FontWeight Bold -Alignment left -BackGroundColor LimeGreen
                                New-HTMLTableHeader -Names "DirectoryCount" -Title $SummaryNotUsedFoldersDirectoryCount -Color White -FontWeight Bold -Alignment left -BackGroundColor LimeGreen
                                New-HTMLTableHeader -Names "Path" -Title " " -ColumnCount 1 -AddRow
                                New-HTMLTableHeader -Names "Size (MB)", "FileCount", "DirectoryCount" -Title "Totals" -Color White -FontWeight Bold -Alignment center -BackGroundColor LimeGreen -AddRow -ColumnCount 3
                            } -PreContent {
                                "<span style='font-size: 1.2em; margin-left: 1em;'>A list of folders that were determined not used under the given path by the searched content objects. It does not include child folders, only `"unique root folders`", so this produces an accurate measurement of capacity used.</span>"
                            }
                        } 
                    }
                }
                $Title = "All not used folders"
                New-HTMLTab -Name $Title {
                    New-HTMLContent -HeaderText $Title {
                        New-HTMLPanel {
                            New-HTMLTable -DataTable $NotUsedFolders -PreContent {
                                "<span style='font-size: 1.2em; margin-left: 1em;'>All folders that were determined not used under the given path by the searched content objects.</span>"
                            }
                        }
                    }
                }
                $Title = "Content objects with invalid path"
                New-HTMLTab -Name $Title {
                    New-HTMLContent -HeaderText $Title {
                        New-HTMLPanel {
                            New-HTMLTable -DataTable ($AllContentObjects.Where( { ($_.SourcePathFlag -eq 3) -And ([string]::IsNullOrEmpty($_.SourcePath) -eq $false) } ) | Select-Object * -ExcludeProperty SourcePathFlag,AllPaths) -PreContent {
                                "<span style='font-size: 1.2em; margin-left: 1em;'>All content objects that have a source path which are not accessible from the computer that ran this script (`"{0}`").</span>" -f $env:COMPUTERNAME
                            }
                        }
                    }
                }
                $Title = "All content objects"
                New-HTMLTab -Name $Title {
                    New-HTMLContent -HeaderText $Title {
                        New-HTMLPanel {
                            New-HTMLTable -DataTable ($AllContentObjects | Select-Object ContentType, UniqueID, Name, SourcePath, SourcePathFlag) -PreContent {
                                "<span style='font-size: 1.2em; margin-left: 1em;'>All searched ConfigMgr content objects.</span>"
                            }
                        }
                    }
                }
            }
            If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Creating HTML report" -Completed -ParentId 1 }
            Write-CMLogEntry -Value "Done creating HTML report" -Severity 1 -Component "Exit" -WriteHost
        }
        catch {
            Write-CMLogEntry -Value ("Failed to create HTML report: {0}" -f $error[0]) -Severity 3 -Component "Exit" -WriteHost
        }
    }

    # Stop clock for runtime
    $StopTime = (Get-Date) - $StartTime

    # Write summary to log
    Write-CMLogEntry -Value ("Content objects: {0}" -f $AllContentObjects.count) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Folders at {0}: {1}" -f $SourcesLocation, $AllFolders.count) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Folders where access denied: {0}" -f ($Result.Where( { $_.UsedBy -like "Access denied*" } ) | Measure-Object | Select-Object -ExpandProperty Count)) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Folders unused: {0}" -f ($NotUsedFolders | Measure-Object | Select-Object -ExpandProperty Count)) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Disk space in `"{0}`" not used by ConfigMgr content objects ({1}): {2} MB" -f $SourcesLocation, ($Commands -replace "Get-CM" -join ", "), $SummaryNotUsedFoldersMB) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Runtime: {0}" -f $StopTime.ToString()) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value "Finished" -Severity 1 -Component "Exit"

    return $Result
}
