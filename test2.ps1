Function Get-LocalPathFromUNCShare {
    param (
        [ValidatePattern("\\\\(.+)(\\).+")]
        [Parameter(Mandatory=$true)]
        [string]$Share
    )

    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)"
    $RegexMatch = [regex]::Match($Share, $Regex)
    $Server = $RegexMatch.Groups[1].Value
    $ShareName = $RegexMatch.Groups[2].Value
    
    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }

    return ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="
}

Function Get-AllSharedFolders {
    Param([String]$Server)
    # Get all shares on server
    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }
    # Iterate through them using hidden PSObject property because $Shares is PSCustomObject
    $AllShares = @{}
    $Shares.PSObject.Properties | Where-Object { $_.TypeNameOfValue -eq "Deserialized.System.String[]" } | ForEach-Object {
        # At this point it's an array
        ForEach ($item in $_) {
            $AllSharesShareName = (($item.Value -match "ShareName") -replace "ShareName=")[0] # There's only ever 1 element in the array
            $AllSharesPath = (($item.Value -match "Path") -replace "Path=")[0] # There's only ever 1 element in the array
            $AllShares += @{ $AllSharesShareName = $AllSharesPath }
        } 
    }
    return $AllShares
}

$Path = "\\SCCM\Applications$\7zip\x64"
$Cache = @{}

$AllPaths = @{}

If ([bool]([System.Uri]$Path).IsUnc -eq $true) {
    
    # Grab server, share and remainder of path in to 4 groups: group 0 (whole match) "\\server\share\folder\folder", group 1 = "server", group 2 = "share", group 3 = "folder\folder"
    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~\\!@#$%^&(){}\'._-]+)*" 
    $RegexMatch = [regex]::Match($Path, $Regex)
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

    If ($Server -as [IPAddress]) {
        $FQDN = [System.Net.Dns]::GetHostEntry("$($Server)").HostName
        $IP = $Server
    }
    Else {
        $FQDN = [System.Net.Dns]::GetHostByName($Server).HostName
        $IP = (((Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)).IPV4Address).IPAddressToString
    }
    $NetBIOS = $FQDN.Split(".")[0]

    If ($Cache.ContainsKey($Server)) {
        # Already have this server's shares cached
    }
    Else {
        # Do not yet have this server's shares cached
        $Cache.Add($Server, (Get-AllSharedFolders -Server $Server))
        $Cache.Add($FQDN, (Get-AllSharedFolders -Server $Server))
        $Cache.Add($IP, (Get-AllSharedFolders -Server $Server))
    }

    # Verify if using UNC drive letter: match only \\server\c$
    $Regex = "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\[a-zA-Z]\$" 
    # We need to take a different approach to building $AllPaths if this is the content object's source path
    If ($Path -match $Regex) {
        # Convert UNC path to local path
        $LocalPath = $Path -replace "^\\\\[a-zA-Z0-9`~!@#$%^&(){}\'._-]+\\" 
        $LocalPath = $LocalPath -replace "\$",":"
        # There was a need to capitalise the driver letter but I can't remember why now
        $LocalPath = [regex]::replace($LocalPath, "^[a-z]:\\", { $args[0].Value.ToUpper() })

        # Start building $AllPaths with what we already know
        $AllPaths.Add($LocalPath, $NetBIOS)
        
        # Now determine all possible paths
        If ($Cache.$Server.count -ge 1) {
            ForEach ($Share in $Cache.$Server.GetEnumerator()) {
                If($LocalPath.StartsWith($Share.Value, "CurrentCultureIgnoreCase")) {
                    $AllPathsArr = @()
                    $AllPathsArr += $LocalPath.replace($Share.Value, "\\$($FQDN)\$($Share.Name)")
                    $AllPathsArr += $LocalPath.replace($Share.Value, "\\$($NetBIOS)\$($Share.Name)")
                    $AllPathsArr += $LocalPath.replace($Share.Value, "\\$($IP)\$($Share.Name)")
                    $AllPathsArr += (("\\" + $FQDN + "\" + $LocalPath) -replace ":", "$")
                    $AllPathsArr += (("\\" + $NetBIOS + "\" + $LocalPath) -replace ":", "$")
                    $AllPathsArr += (("\\" + $IP + "\" + $LocalPath) -replace ":", "$")
                    # This is so we can avoid error messages about dictionary already containing key
                    # Dupes can occur if there are multiple shares within the path
                    ForEach ($item in $AllPathsArr) {
                        If ($AllPaths.ContainsKey($item) -eq $false) {
                            $AllPaths.Add($item, $NetBIOS)
                        }
                    }
                }
            }
        }
        Else {
            $AllPaths.Add($Path, $NetBIOS)
        }
    }
    Else {

        If ($Cache.$Server.ContainsKey($ShareName)) {
            $LocalPath = $Cache.$Server.$ShareName
        }
        # But, what to do if it isn't? ^

        $AllPathsArr = @()
        $AllPathsArr = "\\$($FQDN)\$($ShareName)$($ShareNameRemainder)"
        $AllPathsArr = "\\$($NetBIOS)\$($ShareName)$($ShareNameRemainder)"
        $AllPathsArr = "\\$($IP)\$($ShareName)$($ShareNameRemainder)"
        $AllPathsArr = ("\\$($FQDN)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
        $AllPathsArr = ("\\$($NetBIOS)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
        $AllPathsArr = ("\\$($IP)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
        $AllPathsArr = "$($LocalPath)$($ShareNameRemainder)"

        ForEach ($item in $AllPathsArr) {
            If ($AllPaths.ContainsKey($item) -eq $false) {
                $AllPaths.Add($item, $NetBIOS)
            }
        }

    }
}
Else {
    $AllPaths.Add($Path, $env:COMPUTERNAME)
}

$result = @()
$result += $Cache
$result += $AllPaths

return $result