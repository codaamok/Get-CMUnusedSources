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
        Cody Mathis (Windows Admin slack)
        Chris Kibble (Windows Admin slack)
        Chris Dent (Windows Admin slack)
        PsychoData ("the regex mancer", Windows Admin slack)
        Patrick (Windows Admin slack)
    
    Version history:
    1

    TODO: 
        - $SiteServer should be validated - omg stupid hard
        - Sometimes use $result, sometimes use a name specific exit var in functions - tidy
        - Review comments
        - Output report
        - Use PSDefaultParameter ting
        - Any functions accessing variables in parent scope and not passed as parameter to said function? Clean it!
            - Get-AllFolders for -AltFolderSearch
        - Comment based help
        - Begin Process End blocks, maybe?
        - Decide on component name instead of "Undecided" in main body
        - For the applications, do not see a object in the output for each deployment. Seems like one object for all deployment types with the data of the last retrieved deployment type (maybe)?
        - How can I validate the results?
        - Log is slow to write to at the end
        - End of log file stats e.g. 
        -- Number of content objects (maybe break that down per type)
        -- Number of folders 
        -- Number of access denied
        -- Number of unused folders
        -- Run time?
        - Export posh object to file?
        - Decide on standard for -f format string or inline variable use
        - Decide on stardard for regex > .StartsWith


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
#Requires -Version 5.1
[cmdletbinding(DefaultParameterSetName='1')]
Param (
    [Parameter(Mandatory=$true, Position = 0)]
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
    [Parameter(Mandatory=$true, Position = 1)]
    [ValidatePattern('^[a-zA-Z0-9]{3}$')]
    [string]$SiteCode,
    [Parameter(Mandatory=$true, Position = 2)]
    [ValidateScript({
        If(!(Test-Connection -ComputerName $_ -Count 1 -ErrorAction SilentlyContinue)) {
            Throw "Host `"$($_)`" is unreachable"
        } Else {
            return $true
        }
    })]
    [string]$SiteServer,
    [switch]$Packages,
    [switch]$Applications,
    [switch]$Drivers,
    [switch]$DriverPackages,
    [switch]$OSImages,
    [switch]$OSUpgradeImages,
    [switch]$BootImages,
    [switch]$DeploymentPackages,
    [switch]$AltFolderSearch,
    [switch]$NoProgress,
    [switch]$Log,
    [int32]$LogFileSize = 5MB,
    [int32]$NumOfRotatedLogs = 0,
    [switch]$ObjectExport
)

<#
    Define PSDefaultParameterValues
#>

$JobId = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'

# Write-CMLogEntry
$PSDefaultParameterValues["Write-CMLogEntry:Bias"]=(Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias)
$PSDefaultParameterValues["Write-CMLogEntry:Folder"]=($PSCommandPath | Split-Path -Parent)
$PSDefaultParameterValues["Write-CMLogEntry:FileName"]=(($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + ".log")
$PSDefaultParameterValues["Write-CMLogEntry:Enable"]=$Log.IsPresent
$PSDefaultParameterValues["Write-CMLogEntry:LogFileSize"]=$LogFileSize
$PSDefaultParameterValues["Write-CMLogEntry:MaxNumOfRotatedLogs"]=$NumOfRotatedLogs

<#
    Define functions
#>

Function Write-CMLogEntry {
    # Update with link to Windows Admins github when merged
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
        [int32]$MaxLogFileSize = 5MB,
        [parameter(Mandatory = $false, HelpMessage = 'Maximum number of rotated log files to keep. Set to 0 for unlimited rotated log files.')]
        [ValidateNotNullOrEmpty()]
        [int32]$MaxNumOfRotatedLogs = 0,
        [parameter(Mandatory = $true, HelpMessage = 'A switch that enables the use of this function.')]
        [ValidateNotNullOrEmpty()]
        [switch]$Enable,
        [switch]$WriteHost
    )

    # Runs this regardless of $Enable value
    If ($WriteHost.IsPresent) {
        Write-Host $Value
    }

    If ($Enable) {
        # Determine log file location
        $LogFilePath = Join-Path -Path $Folder -ChildPath $FileName

        If ((([System.IO.FileInfo]$LogFilePath).Exists) -And ($MaxLogFileSize -ne 0)) {

            # Get log size in bytes
            $LogFileSize = [System.IO.FileInfo]$LogFilePath | Select-Object -ExpandProperty Length

            If ($LogFileSize -ge $MaxLogFileSize) {

                # Get log file name without extension
                $LogFileNameWithoutExt = $FileName -replace ([System.IO.Path]::GetExtension($FileName))

                # Get already rolled over logs
                $AllLogs = Get-ChildItem -Path $Folder -Name "$($LogFileNameWithoutExt)_*" -File

                # Sort them numerically (so the oldest is first in the list)
                $AllLogs = $AllLogs | Sort-Object -Descending { $_ -replace '_\d+\.lo_$' }, { [Int]($_ -replace '^.+\d_|\.lo_$') }
            
                ForEach ($Log in $AllLogs) {
                    # Get log number
                    $LogFileNumber = [int32][Regex]::Matches($Log, "_([0-9]+)\.lo_$").Groups[1].Value
                    switch (($LogFileNumber -eq $MaxNumOfRotatedLogs) -And ($MaxNumOfRotatedLogs -ne 0)) {
                        $true {
                            # Delete log if it breaches $MaxNumOfRotatedLogs parameter value
                            [System.IO.File]::Delete("$($Folder)\$($Log)")
                        }
                        $false {
                            # Rename log to +1
                            $NewFileName = $Log -replace "_([0-9]+)\.lo_$","_$($LogFileNumber+1).lo_"
                            [System.IO.File]::Copy("$($Folder)\$($Log)", "$($Folder)\$($NewFileName)", $true)
                        }
                    }
                }

                # Copy main log to _1.lo_
                [System.IO.File]::Copy($LogFilePath, "$($Folder)\$($LogFileNameWithoutExt)_1.lo_", $true)

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
            Write-Warning -Message "Unable to append log entry to $FileName file. Error message: $($_.Exception.Message)"
        }
    }
}

Function Get-CMContent {
    Param(
        [array]$Commands,
        [string]$SiteServer,
        [string]$SiteCode
    )
    [System.Collections.ArrayList]$AllContent = @()
    [hashtable]$ShareCache = @{}
    ForEach ($Command in $Commands) {
        Write-CMLogEntry -Value "Getting: $($Command -replace 'Get-CM')" -Severity 1 -Component "GatherContentObjects"
        $Command = $Command + " | Where-Object SourceSite -eq `"$($SiteCode)`""
        ForEach ($item in (Invoke-Expression $Command)) {
            switch -regex ($Command) {
                "^Get-CMApplication.+" {
                    $AppMgmt = [xml]$item.SDMPackageXML | Select-Object -ExpandProperty AppMgmtDigest
                    $AppName = $AppMgmt.Application.DisplayInfo.FirstChild.Title
                    ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
                        $SourcePaths = $DeploymentType.Installer.Contents.Content.Location
                        # Using ForEach-Object because even if $SourcePaths is null, it will iterate null once which is ideal here where deployment types can have no source path.
                        # Also, A deployment type can have more than 1 source path: for install and uninstall paths
                        $SourcePaths | ForEach-Object {
                            $GetAllPathsResult = Get-AllPaths -Path $_ -Cache $ShareCache -SiteServer $SiteServer
                            $obj = New-Object PSObject
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value "Application"
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value ($DeploymentType | Select-Object -ExpandProperty LogicalName)
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value "$($item.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath "$_"
                            Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $GetAllPathsResult[1]
                            $AllContent.Add($obj) | Out-Null
                        }
                        $ShareCache = $GetAllPathsResult[0]
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
                }   
            }
            Write-CMLogEntry -Value "$($obj.ContentType) - $($obj.UniqueID) - $($obj.Name) - $($obj.SourcePath) - $($obj.AllPaths.Keys -join ',')" -Severity 1 -Component "GatherContentObjects"
        }
    }
    return $AllContent
}

Function Get-AllPaths {
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
            Write-Warning "Unable to interpret path `"$($Path)`""
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
        Write-Warning "Server `"$($Server)`" is unreachable"
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
            Write-Warning "Could not update cache because could not get shared folders from: `"$($FQDN)`""
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
            Write-Warning "Share `"$($ShareName)`" does not exist on `"$($_)`""
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
            Write-CMLogEntry -Value "Couldn't determine path type for `"$($Paths)`" so might have problems accessing folders that breach MAX_PATH limit" -Severity 2 -Component "GatherFolders"
            Write-Warning "Couldn't determine path type for `"$($Paths)`" so might have problems accessing folders that breach MAX_PATH limit"
        }
    }
    
    If ($AltFolderSearch) {
        [System.Collections.ArrayList]$result = Start-AltFolderSearch -FolderName $Path
    }
    Else {
        try {
            [System.Collections.ArrayList]$result = Get-ChildItem -LiteralPath $Path -Directory -Recurse | Select-Object -ExpandProperty FullName
        }
        catch {
            Write-CMLogEntry -Value "Consider using -AltFolderSearch, quiting..." -Severity 3 -Component "GatherFolders"
            Throw "Consider using -AltFolderSearch"
        }
    }

    $result.Add($Path) | Out-Null
    
    switch ($true) {
        ($Path -match "^\\\\\?\\UNC\\") {
            # Matches if starts with \\?\UNC\
            $result = $result -replace [Regex]::Escape("\\?\UNC\"), "\\"
            break
        }
        ($Path -match "^\\\\\?\\[a-zA-Z]{1}:\\") {
            # Matches if starts with \\?\A:\ (A is just an example drive letter used)
            $result = $result -replace [Regex]::Escape("\\?\"), ""
            break
        }
        default {
            # Perhaps don't terminate, but this is just for testing I guess
            Write-CMLogEntry -Value "Couldn't reset $($Path), quiting..." -Severity 3 -Component "GatherFolders"
            Throw "Couldn't reset $($Path)"
        }
    }
    
    $result = $result | Sort-Object

    return $result
}
Function Start-AltFolderSearch {
    Param([string]$FolderName)

    # This exists, because in testing on some older storage devices Get-ChildItem would return an exception "Not enough quota is available to process this command."
    # FullyQualifiedErrorId: DirIOError,Microsoft.PowerShell.Commands.GetChildItemCommand
    # when hit around 50k collection size. Packet trace and deeper digging yielded some sort of SMBv1 exception about a quota limit being exceeded
    # however issue was not applicable to all smbv1 shares.

    # This workaround exists because thanks to Chris Kibble, found a way to recursively grab all folders without hitting said limit.

    # Annoyingly, Get-ChildItem with forced output to an arry @(Get-ChildItem ...) can return an explicit
    # $null value for folders with no subfolders, causing the for loop to indefinitely iterate through
    # working dir when it reaches a null value, so ? $_ -ne $null is needed
    [System.Collections.ArrayList]$FolderList = @(Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName | Where-Object { [string]::IsNullOrEmpty($_) -eq $false })

    ForEach($Folder in $FolderList) {
        $FolderList.Add($(Start-AltFolderSearch -FolderName $Folder)) | Out-Null
    }

    return $FolderList
}

Function Check-FileSystemAccess {
    param
    (
        [string]$Path,
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    # Thanks to Patrick in Windows Admins slack

    [System.Security.Principal.WindowsIdentity]$currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    if (Test-Path $Path)
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
        If((Get-CMSite -SiteCode $SiteCode | Select-Object -ExpandProperty SiteCode) -ne $SiteCode) { throw }

    } catch {
        If((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -ne $null) {
            Set-Location $Path
            Remove-PSDrive -Name $SiteCode -Force
        }
        Throw "Failed to create New-PSDrive with site code `"$($SiteCode)`" and server `"$($Server)`""
    }

}

Write-CMLogEntry -Value "Starting" -Severity 1 -Component "Initilisation"
$StartTime = Get-Date

# Write all parameters passed to script to log
ForEach($var in $PSBoundParameters.GetEnumerator()) {
    Write-CMLogEntry -Value "- $($var.Key): $($var.Value)" -Severity 1 -Component "Initilisation"
}

If ((([System.Uri]$SourcesLocation).IsUnc -eq $false) -And ($env:COMPUTERNAME -ne $SiteServer)) {
    # If user has given local path for $SourcesLocation, need to ensure
    # we don't produce false positives where a similar folder structure exists
    # on the remote machine and site server. e.g. packages let you specify local path
    # on site server
    Write-CMLogEntry -Value "Won't be able to determine unused folders with given local path while running remotely from site server, quitting" -Severity 2 -Component "Initilisation"
    Throw "Will not be able to determine unused folders using local path remote from site server"
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
    $FQDN = [System.Net.Dns]::GetHostEntry("$($SiteServer)") | Select-Object -ExpandProperty HostName
}
Else {
    $FQDN = [System.Net.Dns]::GetHostByName($SiteServer) | Select-Object -ExpandProperty HostName
}
$SiteServer = $FQDN.Split(".")[0]

Write-CMLogEntry -Value "Gathering folders: $($SourcesLocation)" -Severity 1 -Component "GatherFolders"
If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 0 -Status "Calculating number of folders" }
$AllFolders = Get-AllFolders -Path $SourcesLocation #-AltFolderSearch $AltFolderSearch.IsPresent
Write-CMLogEntry -Value "Number of folders: $($AllFolders.count)" -Severity 1 -Component "GatherFolders"

$OriginalPath = Get-Location | Select-Object -ExpandProperty Path
Set-CMDrive -SiteCode $SiteCode -Server $SiteServer -Path $OriginalPath

Write-CMLogEntry -Value "Gathering content objects: $($Commands -replace 'Get-CM')" -Severity 1 -Component "GatherContentObjects"
If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 33 -Status "Getting all CM content objects" }
$AllContentObjects = Get-CMContent -Commands $Commands -SiteServer $SiteServer -SiteCode $SiteCode
Write-CMLogEntry -Value "Number of content objects: $($AllContentObjects.count)" -Severity 1 -Component "GatherContentObjects"

Set-Location $OriginalPath

[System.Collections.ArrayList]$Result = @()

$AllFolders | ForEach-Object -Begin {

    If ($NoProgress -eq $false) { Write-Progress -Id 1 -Activity "Running Get-CMUnusedSources" -PercentComplete 66 -Status "Determinig unused folders" }
    Write-CMLogEntry -Value "Determinig unused folders" -Severity 1 -Component "Undecided"
    
    $NumOfFolders = $AllFolders.count

    # Forcing int data type because double/float for benefit of modulo write-progoress
    If ($NumOfFolders -ge 150) { [int]$FolderInterval = $NumOfFolders * 0.01 } else { $FolderInterval = 2 }
    $FolderCounter = 0

} -Process {

    $FolderCounter++
    $Folder = $_

    If (($FolderCounter % $FolderInterval) -eq 0) { 
        [int]$Percentage = ($FolderCounter / $NumOfFolders * 100)
        If ($NoProgress -eq $false ) { Write-Progress -Id 2 -Activity "Looping through folders in $($SourcesLocation)" -PercentComplete $Percentage -Status "$($Percentage)% complete" -ParentId 1 }
    }
    
    $obj = New-Object PSCustomObject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Folder -Value $Folder

    [System.Collections.ArrayList]$UsedBy = @()
    $IntermediatePath = $false
    $ToSkip = $false
    $NotUsed = $false

    If ((Check-FileSystemAccess -Path $Folder -Rights Read) -ne $true) {
        $UsedBy.Add("Access denied") | Out-Null
        # Still continue anyway because we can still determine if it's an exact match or intermediate path of a content object
    }

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
                    # Package is local host to the site server
                    $UsedBy.Add($ContentObject.Name) | Out-Null
                    break
                }
                (($ContentObject.AllPaths.Keys -contains $Folder) -eq $true) {
                    # By default the ContainsKey method ignores case
                    $UsedBy.Add($ContentObject.Name) | Out-Null
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
    
        $Result.Add($obj) | Out-Null

    }
} -End {

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
        Write-CMLogEntry -Value ($item.Folder + ": " + $item.UsedBy) -Severity $Severity -Component "Undecided"
    }

    # Export $Result to file
    If ($ObjectExport.IsPresent) {
        try {
            Write-CMLogEntry -Value "Exporting object PowerShell object" -Severity 1 -Component "Exit"
            Export-Clixml -LiteralPath (($PSCommandPath | Split-Path -Parent) + "\" + ($PSCommandPath | Split-Path -Leaf) + "_" + $JobId + ".xml") -InputObject $Result
        }
        catch {
            Write-CMLogEntry -Value ("Failed to export PowerShell object: {0}" -f $error[0].Exception.Message) -Severity 3 -Component "Exit"
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