Function Get-AllSharedFolders {
    Param([String]$Server)

    $AllShares = @{}

    try {
        $Shares = Get-WmiObject -ComputerName $Server -Class Win32_Share -ErrorAction Stop | Where-Object {-not [string]::IsNullOrEmpty($_.Path)}
        ForEach ($Share in $Shares) {
            $AllShares += @{ $Share.Name = $Share.Path.TrimEnd("\") }
        }
    }
    catch {
        Write-Warning "Unable to get shared folders from $($Server)"
        $AllShares = $null
    }

    return $AllShares
}

#$Path = "\\sccm\Applications$\7zip\x64"

$Cache = @{}
$SCCMServer = "SCCM"

$AllPaths = @{}
$result = @()

Clear-Variable -Name "Server"
Clear-Variable -Name "NetBIOS"
Clear-Variable -Name "IP"
Clear-Variable -Name "FQDN"
Clear-Variable -Name "PathType"

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
        $Allpaths.Add($Path, $SCCMServer)
        $result += $Cache
        $result += $AllPaths
        return $result
    }
    default { 
        Write-Warning "Unable to interpret path: `"$($Path)`""
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
    Write-Warning "Server `"$($Server)`" is unreachable"
    $AllPaths.Add($Path, $null)
    $result += $Cache
    $result += $AllPaths
    return $result
}

##### Update the cache of shared folders and their local paths

If ($Cache.ContainsKey($Server) -eq $false) {
    # Do not yet have this server's shares cached
    # $AllSharedFolders is null if couldn't connect to serverr to get all shared folders
    $NetBIOS,$FQDN,$IP | Where-Object { $_ -ne $null } | ForEach-Object {
        $AllSharedFolders = Get-AllSharedFolders -Server $Server
        If ($AllSharedFolders -ne $null) {
            $Cache.Add($_, $AllSharedFolders)
        }
        Else {
            Write-Warning "Could not update cache because could not get shared folders from: `"$($Server)`" / `"$($_)`""
        }
    }
}

##### Build the AllPaths property

$AllPathsArr = @()

$NetBIOS,$FQDN,$IP | Where-Object { $_ -ne $null } | ForEach-Object -Process {
    If ($Cache.$_.ContainsKey($ShareName)) {
        $LocalPath = $Cache.$_.$ShareName
        $AllPathsArr += ("\\$($_)\$($LocalPath)$($ShareRemainder)" -replace ':', '$')
    }
    Else {
        Write-Warning "Share `"$($ShareName)`" does not exist on `"$($_)`""
    }
    $AllPathsArr += "\\$($_)\$($ShareName)$($ShareRemainder)"
} -End {
    If ($LocalPath -ne $null) {
        If ($LocalPath -match "^[a-zA-Z]:$") {
            $AllPathsArr += "$($LocalPath)\"
        }
        Else {
            $AllPathsArr += "$($LocalPath)$($ShareRemainder)"
        }
    }
}

ForEach ($item in $AllPathsArr) {
    If ($AllPaths.ContainsKey($item) -eq $false) {
        $AllPaths.Add($item, $NetBIOS)
    }
}

$result += $Cache
$result += $AllPaths
$result[1].GetEnumerator().Name | ForEach-Object { "$(Test-Path -Path $_) " + $_ }