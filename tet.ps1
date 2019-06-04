Function Get-AllSharedFolders {
    Param([String]$Server)

    $AllShares = @{}

    try {
        $Shares = Get-WmiObject -ComputerName $Server -Class Win32_Share -ErrorAction Stop | Where-Object {-not [string]::IsNullOrEmpty($_.Path)}
        ForEach ($Share in $Shares) {
            # The TrimEnd method is only really concerned for drive letter shares
            # as they're usually stored as f$ = "F:\" and this messes up Get-AllPaths a little
            $AllShares += @{ $Share.Name = $Share.Path.TrimEnd("\") }
        }
    }
    catch {
        $AllShares = $null
    }

    return $AllShares
}

Function Get-AllPaths {
    # In this function you'll see use of $Hashtable.Keys -contains "value" as opposed to the $Hashtable.ContainsKey($value) method
    # because if $value is null while using ContainsKey then a non-terminating error is printed
    param (
        [string]$Path,
        [hashtable]$Cache,
        [string]$SCCMServer
    )

    $AllPaths = @{}
    $result = @()

    If ([string]::IsNullOrEmpty($Path) -eq $false) {
        $Path = $Path.TrimEnd("\")
    }

    ##### Determine path type

    switch ($true) {
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)(\\[a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)") {
            # Path that is \\server\share\folder
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$Matches[3]
        }
        ($Path -match "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)$") {
            # Path that is \\server\share
            $Server,$ShareName,$ShareRemainder = $Matches[1],$Matches[2],$null
        }
        ($Path -match "^[a-zA-Z]:\\") {
            # Path that is just drive letter
            $AllPaths.Add($Path, $SCCMServer)
            $result += $Cache
            $result += $AllPaths
            return $result
        }
        ([string]::IsNullOrEmpty($Path) -eq $true) {
            $result += $Cache
            $result += $AllPaths
            return $result
        }
        default { 
            Write-Warning "Unable to interpret path `"$($Path)`", used by `"$($obj.Name)`""
            $AllPaths.Add($Path, $null)
            $result += $Cache
            $result += $AllPaths
            return $result
        }
    }

    ##### Determine FQDN, IP and NetBIOS

    If (Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue) {
        If ($Server -as [IPAddress]) {
            try {
                $FQDN = [System.Net.Dns]::GetHostEntry($Server).HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                $FQDN = $null
            }
            $IP = $Server
        }
        Else {
            try {
                $FQDN = [System.Net.Dns]::GetHostByName($Server).HostName
                $NetBIOS = $FQDN.Split(".")[0]
            }
            catch {
                $FQDN = $null
            }
            $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
        }
    }
    Else {
        Write-Warning "Server `"$($Server)`" is unreachable, used by `"$($obj.Name)`""
        $AllPaths.Add($Path, $null)
        $result += $Cache
        $result += $AllPaths
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
            Write-Warning "Could not update cache because could not get shared folders from: `"$($FQDN)`", used by `"$($obj.Name)`""
        }
    }

    ##### Build the AllPaths property

    $AllPathsArr = @()

    ## Build AllPaths based on share name from given Path

    $NetBIOS,$FQDN,$IP | Where-Object { [string]::IsNullOrEmpty($_) -eq $false } | ForEach-Object -Process {
        $AltServer = $_
        $LocalPath = ($Cache.$AltServer.GetEnumerator() | Where-Object { $_.Key -eq $ShareName }).Value
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            $AllPathsArr += ("\\$($AltServer)\$($LocalPath)$($ShareRemainder)" -replace ':', '$')
            $SharesWithSamePath = ($Cache.$AltServer.GetEnumerator() | Where-Object { $_.Value -eq $LocalPath }).Key
            $SharesWithSamePath | ForEach-Object {
                $AltShareName = $_
                $AllPathsArr += "\\$($AltServer)\$($AltShareName)$($ShareRemainder)"
            }
        }
        Else {
            Write-Warning "Share `"$($ShareName)`" does not exist on `"$($_)`", used by `"$($obj.Name)`""
        }
        $AllPathsArr += "\\$($AltServer)\$($ShareName)$($ShareRemainder)"
    } -End {
        If ([string]::IsNullOrEmpty($LocalPath) -eq $false) {
            If ($LocalPath -match "^[a-zA-Z]:$") {
                $AllPathsArr += "$($LocalPath)\"
            }
            Else {
                $AllPathsArr += "$($LocalPath)$($ShareRemainder)"
            }
        }
    }

    ForEach ($item in $AllPathsArr) {
        If (($AllPaths.Keys -contains $item) -eq $false) {
            $AllPaths.Add($item, $NetBIOS)
        }
    }

    $result += $Cache
    $result += $AllPaths
    return $result
}

$Cache = @{}
Get-AllPaths -Path "\\sccm\Applications$\7zip" -Cache $Cache -SCCMServer "SCCM"