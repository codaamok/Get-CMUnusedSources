<#
.SYNOPSIS
Get-CMUnusedSources will tell you what folders are not used by ConfigMgr in a given path.

.DESCRIPTION
Check out https://www.cookadam.co.uk/get-cmunusedsources and https://github.com/codaamok/Get-CMUnusedSources.

.PARAMETER SourcesLocation
The path to the directory you store your ConfigMgr sources. Can be a UNC or local path.
Must be a valid path that you have read access to.

.PARAMETER SiteCode
The site code of the ConfigMgr site you wish to query for content objects.

.PARAMETER SiteServer
The site server of the given ConfigMgr site code.
The server must be reachable over a network.

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
Specify this to enable logging.
The log file(s) will be saved to the same directory as this script with a name of <scriptname>_<datetime>.log. Rolled log files will follow a naming convention of <filename>_1.lo_ where the int increases for each rotation.

.PARAMETER LogFileSize
Set the maximum size you want for each rolled over log file. This is only applicable if NumOfRotatedLogs is greater than 0.
Default value is 5MB. The unit of measurement is bytes however you can specify units such as KB, MB etc.

.PARAMETER NumOfRotatedLogs
Set the maximum number of log files you wish to keep.
Default value is 5MB. Specify 0 for unlimited.

.PARAMETER ObjectExport
Specify this option if you wish to export the PowerShell result object to an XML file.
The XML file be saved to the same directory as this script with a name of <scriptname>_<datetime>.xml. It can easily be reimported using Import-Clixml cmdlet.

.PARAMETER HtmlReport
Specify this option to enable the generation for a HTML report of the result. Doing this will force you to have the PSWriteHtml module installed.
For more information on PSWriteHTML: https://github.com/EvotecIT/PSWriteHTML
The HTML file will be saved to the same directory as this script with a name of <scriptname>_<datetime>.html.

.PARAMETER Threads
Set the number of threads you wish to use for concurrent processing of this script.
Default value is number of processes from env var NUMBER_OF_PROCESSORS minus 1. 
A validation is in place to prevent you from going above NUMBER_OF_PROCESSORS+1. If you want to do that, you can. Remove the ValidateScript block on the parameter. I've put this place just in case you you run this on your already busy site server.

.INPUTS

.EXAMPLE

C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation \\sccm\Applications$ -SiteCode ACC -SiteServer SCCM -Applications -Log -LogFileSize 10MB -NumOfRotatedLogs 5 -ObjectExport -HtmlReport -Threads 2

.EXAMPLE

C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation F:\ -SiteCode ACC -SiteServer SCCM -Log -HtmlReport

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
    [Parameter(Mandatory=$false, HelpMessage="Maximum size per log file.")]
    [ValidatePattern("^(?i)[0-9]+(mb|kb|gb)?$")]
    [int32]$LogFileSize = 5MB,
    [Parameter(Mandatory=$false, HelpMessage="Maximum number of rotated log files to keep.")]
    [int32]$NumOfRotatedLogs = 0,
    [Parameter(Mandatory=$false, HelpMessage="Generate XML export of PowerShell object with the result.")]
    [switch]$ObjectExport,
    [Parameter(Mandatory=$false, HelpMessage="Generate HTML report of the result.")]
    [switch]$HtmlReport,
    [Parameter(Mandatory=$false, HelpMessage="Number of threads to use for execution.")]
    [ValidateScript({
        If($_ -gt ($env:NUMBER_OF_PROCESSORS+1)) {
            throw "If you really want to use this many threads, you'll have to modify the script yourself to allow you."
        } Else {
            return $true
        }
    })]
    [int32]$Threads = [int]$env:NUMBER_OF_PROCESSORS-1
)

<#
TODO: 
        - $SiteServer should be validated - omg stupid hard
        - Review comments
        - Output report
        - Use PSDefaultParameter ting
        - Any functions accessing variables in parent scope and not passed as parameter to said function? Clean it!
            - Get-AllFolders for -AltFolderSearch
        - Begin Process End blocks, maybe?
        - How can I validate the results?
        - Can get-allpaths use Join-Path in any way?


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

    What to put in report:
        - Total space consumed by unused folders (robocopy?)
        - Content objects with invalid path
        - All unused folders

    Final remarks:
        - The given ID in the log/report for DeploymentTypes unfortunately isn't what you need for the Get-CMDeploymentType cmdlet. However you can use the names. The XML for the deployment types does not contain the CI_ID, which is what used by Get-CMDeploymentType.
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
$PSDefaultParameterValues["Write-CMLogEntry:MaxLogFileSize"]=$LogFileSize
$PSDefaultParameterValues["Write-CMLogEntry:MaxNumOfRotatedLogs"]=$NumOfRotatedLogs

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
    <#
    .SYNOPSIS
    Get all ConfigMgr objects that can hold content, i.e. content objects.
    .DESCRIPTION
    Using the ConfigMgr PoSH cmdlets, in the $Commands array, get all content objects and filter them to the given site code. 
    For each content object, create a PSCustomObject with the needed properties and return them all in an ArrayList collection.
    Called by main body.
    .OUTPUTS
    System.Object.ArrayList where each element will be a content object of System.Object.PSCustomObject.
    #>
    Param(
        [array]$Commands,
        [string]$SiteServer,
        [string]$SiteCode
    )
    [System.Collections.ArrayList]$AllContent = @()
    [hashtable]$ShareCache = @{}
    ForEach ($Command in $Commands) {
        Write-CMLogEntry -Value ("Getting: {0}" -f $Command -replace "Get-CM") -Severity 1 -Component "GatherContentObjects"
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
                            $GetAllPathsResult = Get-AllPaths -Path $_ -Cache $ShareCache -SiteServer $SiteServer
                            $obj = New-Object PSObject
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Application"
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value ($DeploymentType | Select-Object -ExpandProperty LogicalName)
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value ("{0}::{1}" -f $item.LocalizedDisplayName,$DeploymentType.Title.InnerText)
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $_
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                            $AllContent.Add($obj) | Out-Null
                        }
                        $ShareCache = $GetAllPathsResult[0]
                        Write-CMLogEntry -Value ("{0} - {1} - {2} - {3} - {4}" -f $obj.ContentType,$obj.UniqueID,$obj.Name,$obj.SourcePath,$obj.AllPaths.Keys -join ",") -Severity 1 -Component "GatherContentObjects"
                    }
                }
                "^Get-CMDriver\s.+" { # I don't actually think it's possible for a driver to not have source path set
                    $SourcePath = $item.ContentSourcePath
                    $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SiteServer $SiteServer    
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Driver"
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.CI_ID
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.LocalizedDisplayName
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                    $AllContent.Add($obj) | Out-Null
                    $ShareCache = $GetAllPathsResult[0]
                    Write-CMLogEntry -Value ("{0} - {1} - {2} - {3} - {4}" -f $obj.ContentType,$obj.UniqueID,$obj.Name,$obj.SourcePath,$obj.AllPaths.Keys -join ",") -Severity 1 -Component "GatherContentObjects"
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
                    $GetAllPathsResult = Get-AllPaths -Path $SourcePath -Cache $ShareCache -SiteServer $SiteServer
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value $ContentType
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.PackageId
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.Name
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath $SourcePath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                    $AllContent.Add($obj) | Out-Null
                    $ShareCache = $GetAllPathsResult[0]
                    Write-CMLogEntry -Value ("{0} - {1} - {2} - {3} - {4}" -f $obj.ContentType,$obj.UniqueID,$obj.Name,$obj.SourcePath,$obj.AllPaths.Keys -join ",") -Severity 1 -Component "GatherContentObjects"
                }   
            }
        }
    }
    return $AllContent
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
    System.Object.ArrayList with always only two elements; $AllPaths (hashtable, the calculated list of "all paths" for the given $Path), and $Cache (hashtable, the shared folders cache).
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
        \\sccm\BackOfDicks$            sccm
        \\sccm.acc.local\F$\Applica... sccm
        \\192.168.175.11\BackOfDicks$  sccm
        F:\Applications                sccm
        \\sccm.acc.local\BackOfDicks$  sccm
        \\sccm\F$\Applications         sccm
    #>
    param (
        [string]$Path,
        [hashtable]$Cache,
        [string]$SiteServer
    )

    [System.Collections.ArrayList]$result = @()
    [hashtable]$AllPaths = @{}

    If ([string]::IsNullOrEmpty($Path) -eq $false) {
        $Path = $Path.TrimEnd("\")
    }

    ##### Determine path type

    switch ($true) {
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)(\\[a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)") {
            # Path that is \\server\share\folder
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$Matches[3]
            break
        }
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)$") {
            # Path that is \\server\share
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$null
            break
        }
        ($Path -match "^[a-zA-Z]:\\") {
            # Path that is local
            $AllPaths.Add($Path, $SiteServer)
            $result.Add($Cache) | Out-Null
            $result.Add($AllPaths) | Out-Null
            return $result
        }
        ([string]::IsNullOrEmpty($Path) -eq $true) {
            $result.Add($Cache) | Out-Null
            $result.Add($AllPaths) | Out-Null
            return $result
        }
        default { 
            $Message = "Unable to interpret path `"{0}`"" -f $Path
            Write-Warning $Message
            Write-CMLogEntry -Value $Message -Severity 2 -Component "GatherContentObjects"
            $AllPaths.Add($Path, $null)
            $result.Add($Cache) | Out-Null
            $result.Add($AllPaths) | Out-Null
            return $result
        }
    }

    ##### Determine FQDN, IP and NetBIOS

    If (Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue) {
        If ($Server -as [IPAddress]) {
            try {
                $FQDN = [System.Net.Dns]::GetHostEntry($Server) | Select-Object -ExpandProperty HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                $FQDN = $null
            }
            $IP = $Server
        }
        Else {
            try {
                $FQDN = [System.Net.Dns]::GetHostByName($Server) | Select-Object -ExpandProperty HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                $FQDN = $null
            }
            $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
        }
    }
    Else {
        $Message = "Server `"{0}`" is unreachable" -f $Server
        Write-Warning $Message
        Write-CMLogEntry -Value $Message -Severity 2 -Component "GatheringContentObjects"
        $AllPaths.Add($Path, $null)
        $result.Add($Cache) | Out-Null
        $result.Add($AllPaths) | Out-Null
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
            $Message = "Could not update cache because could not get shared folders from: `"{0}`"" -f $FQDN
            Write-Warning $Message
            Write-CMLogEntry -Value $Message -Severity 2 -Component "GatheringContentObjects"
        }
    }

    ##### Build the AllPaths property

    [System.Collections.ArrayList]$AllPathsArr = @()

    ## Build AllPaths based on share name from given Path

    $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object -Process {
        $AltServer = $_
        $LocalPath = $Cache.$AltServer.GetEnumerator() | Where-Object { $_.Key -eq $ShareName } | Select-Object -ExpandProperty Value
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            $AllPathsArr.Add(("\\$($AltServer)\$($LocalPath)$($ShareRemainder)" -replace ':', '$')) | Out-Null
            $SharesWithSamePath = $Cache.$AltServer.GetEnumerator() | Where-Object { $_.Value -eq $LocalPath } | Select-Object -ExpandProperty Key
            $SharesWithSamePath | ForEach-Object -Process {
                $AltShareName = $_
                $AllPathsArr.Add("\\$($AltServer)\$($AltShareName)$($ShareRemainder)") | Out-Null
            }
        }
        Else {
            $Message = "Share `"{0}`" does not exist on `"{1}`"" -f $ShareName,$_
            Write-Warning $Message
            Write-CMLogEntry -Value $Message -Severity 2 -Component "GatheringContentObjects"
        }
        $AllPathsArr.Add("\\$($AltServer)\$($ShareName)$($ShareRemainder)") | Out-Null
    } -End {
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            If ($LocalPath -match "^[a-zA-Z]:$") {
                # Match if drive letter
                $AllPathsArr.Add("$($LocalPath)\") | Out-Null
            }
            Else {
                $AllPathsArr.Add("$($LocalPath)$($ShareRemainder)") | Out-Null
            }
        }
    }

    ForEach ($item in $AllPathsArr) {
        If (($AllPaths.Keys -contains $item) -eq $false) {
            $AllPaths.Add($item, $NetBIOS)
        }
    }

    $result.Add($Cache) | Out-Null
    $result.Add($AllPaths) | Out-Null
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
        $Shares = Get-WmiObject -ComputerName $Server -Class Win32_Share -ErrorAction Stop | Where-Object {-not [string]::IsNullOrEmpty($_.Path)}
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
    System.Object.ArrayList of folder full names.
    #>
    Param(
        [string]$Path,
        [bool]$AltFolderSearch
    )

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
            $Message = "Couldn't determine path type for `"{0}`" so might have problems accessing folders that breach MAX_PATH limit" -f $Paths
            Write-Warning $Message
            Write-CMLogEntry -Value $Message -Severity 2 -Component "GatherFolders"
        }
    }
    
    If ($AltFolderSearch) {
        [System.Collections.ArrayList]$Folders = Start-AltFolderSearch -FolderName $Path
    }
    Else {
        try {
            [System.Collections.ArrayList]$Folders = Get-ChildItem -LiteralPath $Path -Directory -Recurse | Select-Object -ExpandProperty FullName
        }
        catch {
            Write-CMLogEntry -Value "Consider using -AltFolderSearch, quiting..." -Severity 3 -Component "GatherFolders"
            throw "Consider using -AltFolderSearch"
        }
    }

    $Folders.Add($Path) | Out-Null
    
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
            # Perhaps don't terminate, but this is just for testing I guess
            $Message = "Couldn't reset {0}" -f $Path
            Write-CMLogEntry -Value $Message -Severity 3 -Component "GatherFolders"
            throw $Message
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
    System.Object.ArrayList of folder full names.
    #>
    Param([string]$FolderName)

    # Annoyingly, Get-ChildItem with forced output to an arry @(Get-ChildItem ...) can return an explicit
    # $null value for folders with no subfolders, causing the for loop to indefinitely iterate through
    # working dir when it reaches a null value, so is null check is needed
    [System.Collections.ArrayList]$Folders = @(Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName | Where-Object { [string]::IsNullOrEmpty($_) -eq $false })

    ForEach($Folder in $Folders) {
        $Folders.Add($(Start-AltFolderSearch -FolderName $Folder)) | Out-Null
    }

    return $Folders
}

Function Get-FileSystemAccess {
    <#
    .SYNOPSIS
    Check for read access on a given folder.
    .DESCRIPTION
    This is a very fast method of checking for read access on $Path by pulling access rules and comparing it to the ID to the user's context running this cricket.
    I can not take any credit for this function. Huge thanks to Patrick in WinAdmins Slack!
    Called by main body.
    .OUTPUTS
    Boolean
    #>
    param
    (
        [string]$Path,
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    # Thanks to Patrick in Windows Admins slack

    [System.Security.Principal.WindowsIdentity]$currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    if ([System.IO.Directory]::Exists($Path))
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
    <#
    .SYNOPSIS
    Import ConfigMgr module, create ConfigrMgr PS drive and set location to it.
    .DESCRIPTION
    Set current working directory to site code for access to ConfigMgr cmdlets. Some validation is in place to verify the site code marrys up to be of $Server.
    Called by main body.
    .OUTPUTS
    Boolean
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
            throw "Failed to import Configuration Manager module"
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
        throw "Failed to create New-PSDrive with site code `"{0}`" and server `"{1}`"" -f $SiteCode, $Server
    }

}

Write-CMLogEntry -Value "Starting" -Severity 1 -Component "Initilisation" -WriteHost
$StartTime = Get-Date

# Write all parameters passed to script to log
ForEach($item in $PSBoundParameters.GetEnumerator()) {
    Write-CMLogEntry -Value ("- {0}: {1}" -f $item.Key, $item.Value) -Severity 1 -Component "Initilisation"
}

If ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($env:COMPUTERNAME -ne $SiteServer)) {
    # If user has given local path for $SourcesLocation, need to ensure
    # we don't produce false positives where a similar folder structure exists
    # on the remote machine and site server. e.g. packages let you specify local path
    # on site server
    Write-CMLogEntry -Value "Won't be able to determine unused folders with given local path while running remotely from site server, quitting" -Severity 2 -Component "Initilisation" -WriteHost
    throw "Will not be able to determine unused folders using local path remote from site server"
}

# Import PSWriteHtml module report if -HtmlReport is present
If ($HtmlReport.IsPresent -eq $true) {
    try {
        Import-Module PSWriteHTML -ErrorAction Stop
    }
    catch {
        $Message = "Unable to import PSWriteHtml module: {0}" -f $error[0].Exception.Message
        Write-CMLogEntry -Value $Message -Severity 3 -Component "Initilsation" -WriteHost
        throw $Message
    }
}

[System.Collections.ArrayList]$AllContentObjects = @()

switch ($true) {
    ($Packages.IsPresent) {
        [array]$Commands += "Get-CMPackage"
    }
    ($Applications.IsPresent) {
        [array]$Commands += "Get-CMApplication"
    }
    ($Drivers.IsPresent) {
        [array]$Commands += "Get-CMDriver"
    }
    ($DriverPackages.IsPresent) {
        [array]$Commands += "Get-CMDriverPackage"
    }
    ($OSImages.IsPresent) {
        [array]$Commands += "Get-CMOperatingSystemImage"
    }
    ($OSUpgradeImages.IsPresent) {
        [array]$Commands += "Get-CMOperatingSystemInstaller"
    }
    ($BootImages.IsPresent) {
        [array]$Commands += "Get-CMBootImage"
    }
    ($DeploymentPackages.IsPresent) {
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

Write-CMLogEntry -Value ("Gathering folders: {0}" -f $SourcesLocation) -Severity 1 -Component "GatherFolders" -WriteHost
If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 0 -Status ("Gathering all folders at: {0}" -f $SourcesLocation) }
$AllFolders = Get-AllFolders -Path $SourcesLocation #-AltFolderSearch $AltFolderSearch.IsPresent
Write-CMLogEntry -Value ("Number of folders: {0}" -f $AllFolders.count) -Severity 1 -Component "GatherFolders" -WriteHost

$OriginalPath = Get-Location | Select-Object -ExpandProperty Path
Set-CMDrive -SiteCode $SiteCode -Server $SiteServer -Path $OriginalPath

Write-CMLogEntry -Value ("Gathering content objects: {0}" -f ($Commands -replace "Get-CM" -join ", ")) -Severity 1 -Component "GatherContentObjects" -WriteHost
If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 33 -Status ("Gathering CM content objects: {0}" -f ($Commands -replace "Get-CM" -join ", ")) }
$AllContentObjects = Get-CMContent -Commands $Commands -SiteServer $SiteServer -SiteCode $SiteCode
Write-CMLogEntry -Value ("Number of content objects: {0}" -f $AllContentObjects.count) -Severity 1 -Component "GatherContentObjects" -WriteHost

Set-Location $OriginalPath

$AllFolders | ForEach-Object -Begin {

    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 66 -Status "Determining unused folders" }
    Write-CMLogEntry -Value ("Determinig unused folders, using {0} threads" -f $Threads) -Severity 1 -Component "Processing" -WriteHost
    
    $Definition = Get-Content Function:\Get-FileSystemAccess -ErrorAction Stop
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'Get-FileSystemAccess', $Definition
    $initialSessionState = [InitialSessionState]::CreateDefault()
    $InitialSessionState.Commands.Add($SessionStateFunction)

    $RSPool = [RunspaceFactory]::CreateRunspacePool(1, $Threads, $InitialSessionState, $Host)
    $RSPool.ApartmentState = "MTA"
    $RSPool.Open()
    $RSResults = @()
    $RSScriptBlock = {
        Param (
            [string]$RSFolder,
            [System.Collections.ArrayList]$RSAllContentObjects
        )

        $obj = New-Object PSCustomObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Folder -Value $RSFolder

        [System.Collections.ArrayList]$UsedBy = @()
        $IntermediatePath = $false
        $ToSkip = $false
        $NotUsed = $false

        If ((Get-FileSystemAccess -Path $RSFolder -Rights Read) -ne $true) {
            $UsedBy.Add("Access denied") | Out-Null
            # Still continue anyway because we can still determine if it's an exact match or intermediate path of a content object
        }

        If ($RSFolder.StartsWith($ToSkip)) {
            # Should probably rename $NotUsed to something more appropriate to truely reflect its meaning
            # This is here so we don't walk through completely unused folder + sub folders
            # Unused folders + sub folders are learnt for each loop of a new folder structure and thus each loop of all content objects
            $NotUsed = $true
        }
        Else {

            ForEach ($ContentObject in $RSAllContentObjects) {

                switch($true) {
                    ([string]::IsNullOrEmpty($ContentObject.SourcePath) -eq $true) {
                        break
                    }
                    ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($ContentObject.AllPaths.($RSFolder) -eq $env:COMPUTERNAME)) {
                        # Package is local host to the site server
                        $UsedBy.Add($ContentObject.Name) | Out-Null
                        break
                    }
                    (($ContentObject.AllPaths.Keys -contains $RSFolder) -eq $true) {
                        # By default the ContainsKey method ignores case
                        $UsedBy.Add($ContentObject.Name) | Out-Null
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
                        $ToSkip = $RSFolder
                        $NotUsed = $true
                    }
                }

            }

            switch ($true) {
                ($UsedBy.count -gt 0) {
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UsedBy -Value (($UsedBy) -join ', ')
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
        }
        return $obj
    }


} -Process {

    $Folder = $_

    $Runspace = [PowerShell]::Create()
    $null = $Runspace.AddScript($RSScriptBlock)
    $null = $Runspace.AddArgument($Folder)
    $null = $Runspace.AddArgument($AllContentObjects)
    $Runspace.Runspacepool = $RSPool
    $RSResults += [PSCustomObject]@{ Pipe = $Runspace; Status = $Runspace.BeginInvoke() }
    
} -End {
    
    [System.Collections.ArrayList]$Result = @()

    # Process runspaces, their results and clean up when complete
    $TotalRunspaces = $RSResults | Measure-Object | Select-Object -ExpandProperty Count
    while ($RSResults.Status -ne $null) {
        If ($NoProgress -eq $false) { 
            $TotalNotComplete = $RSResults | Where-Object { $_.Status -eq $null } | Measure-Object | Select-Object -ExpandProperty Count
            Write-Progress -Id 2 -Activity "Evaluating folders" -Status ("{0} folders remaining" -f ($TotalRunspaces-$TotalNotComplete)) -PercentComplete ($TotalNotComplete/$TotalRunspaces * 100) -ParentId 1
        }
        $Completed = $RSResults | Where-Object { $_.Status.IsCompleted -eq $true }
        ForEach ($item in $Completed) {
            # Reference index 0 so we can grab the PSCustomobject inside the PSDataCollection object
            $Result.Add(($item.Pipe.EndInvoke($item.Status)[0])) | Out-Null
            $item.Pipe.Dispose()
            $item.Status = $null
        }
        Start-Sleep -Seconds 2
    }

    # Clean up runspace pool
    $RSPool.Dispose()

    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 2 -Activity "Evaluating folders" -Completed -ParentId 1 }
    If ($NoProgress.IsPresent -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 100 -Status "Finishing up" }

    # Write $Result to log file
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

    # Export $Result to file
    If ($ObjectExport.IsPresent) {
        try {
            Write-CMLogEntry -Value "Exporting object PowerShell object" -Severity 1 -Component "Exit" -WriteHost
            Export-Clixml -LiteralPath (($PSCommandPath | Split-Path -Parent) + "\" + ($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + ".xml") -InputObject $Result
        }
        catch {
            Write-CMLogEntry -Value ("Failed to export PowerShell object: {0}" -f $error[0].Exception.Message) -Severity 3 -Component "Exit" -WriteHost
        }
    }

    # Write $Result to HTML
    If ($HtmlReport.IsPresent) {
        try {
            Write-CMLogEntry -Value "Creating HTML report" -Severity 1 -Component "Exit" -WriteHost
            New-HTML -TitleText ("Get-CMUnusedSources - {0}" -f $JobId) -UseCssLinks:$true -UseJavaScriptLinks:$true -FilePath (($PSCommandPath | Split-Path -Parent) + "\" + ($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + ".html") -ShowHTML {
                New-HTMLContent -HeaderText "All folders" -CanCollapse {
                    New-HTMLPanel {
                        New-HTMLTable -DataTable $Result -HideFooter 
                    }
                }
            }
        }
        catch {
            Write-CMLogEntry -Value ("Failed to create HTML report: {0}" -f $error[0].Exception.Message) -Severity 3 -Component "Exit" -WriteHost
        }
    }


    # Stop clock for runtime
    $StopTime = (Get-Date) - $StartTime

    # Write summary to log
    Write-CMLogEntry -Value ("Total number of content objects: {0}" -f $AllContentObjects.count) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Total number of folders at {0}: {1}" -f $SourcesLocation, $AllFolders.count) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Total number of folders where access denied: {0}" -f ($Result | Where-Object { $_.UsedBy -like "Access denied*" } | Measure-Object).count) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Total number of folders unused: {0}" -f ($Result | Where-Object {$_.UsedBy -eq "Not used"} | Measure-Object).count) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value ("Total runtime: {0}" -f $StopTime.ToString()) -Severity 1 -Component "Exit" -WriteHost
    Write-CMLogEntry -Value "Finished" -Severity 1 -Component "Exit"

    return $Result
}