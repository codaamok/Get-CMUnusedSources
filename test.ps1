Function Get-AllPaths {
    param (
        [string]$Path,
        [hashtable]$Cache
    )

    $AllPaths = @{}

    If ([bool]([System.Uri]$Path).IsUnc -eq $true) {
        
        # Grab server, share and remainder of path in to 4 groups: group 0 (whole match) "\\server\share\folder\folder", group 1 = "server", group 2 = "share", group 3 = "folder\folder"
        $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)\\([a-zA-Z0-9`~\\!@#$%^&(){}\'._ -]+)*" 
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
            $Cache.Add($NetBIOS, (Get-AllSharedFolders -Server $Server))
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
            $AllPathsArr += "\\$($FQDN)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += "\\$($NetBIOS)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += "\\$($IP)\$($ShareName)$($ShareNameRemainder)"
            $AllPathsArr += ("\\$($FQDN)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
            $AllPathsArr += ("\\$($NetBIOS)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
            $AllPathsArr += ("\\$($IP)\$($LocalPath)$($ShareNameRemainder)" -replace ':', '$')
            $AllPathsArr += "$($LocalPath)$($ShareNameRemainder)"
    
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
}

Function Get-LocalPathFromUNCShare {
    param (
        [ValidatePattern("\\\\(.+)(\\).+")]
        [Parameter(Mandatory=$true)]
        [string]$Share
    )

    $Regex = "^\\\\([a-zA-Z0-9`~!@#$%^&(){}\'._-]+)\\([a-zA-Z0-9`~!@#$%^&(){}\'._ -]+)"
    $RegexMatch = [regex]::Match($Share, $Regex)
    $Server = $RegexMatch.Groups[1].Value
    $ShareName = $RegexMatch.Groups[2].Value
    
    $Shares = Invoke-Command -ComputerName $Server -ScriptBlock { get-itemproperty -path registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares }

    return ($Shares.$ShareName | Where-Object {$_ -match 'Path'}) -replace "Path="
}

Set-Location ACC:

#$Commands = "Get-CMPackage", "Get-CMDriverPackage", "Get-CMBootImage", "Get-CMOperatingSystemImage", "Get-CMOperatingSystemInstaller", "Get-CMSoftwareUpdateDeploymentPackage", "Get-CMApplication", "Get-CMDriver"
$Commands = "Get-CMApplication"
$AllContent = @()
$ShareCache = @{}
ForEach ($Command in $Commands) {
    ForEach ($item in (Invoke-Expression $Command)) {
        switch ($Command) {
            "Get-CMApplication" {
                $AppMgmt = ([xml]$item.SDMPackageXML).AppMgmtDigest
                $AppName = $AppMgmt.Application.DisplayInfo.FirstChild.Title
                ForEach ($DeploymentType in $AppMgmt.DeploymentType) {
                    $obj = New-Object PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value "$($DeploymentType.AuthoringScopeId)/$($DeploymentType.LogicalName)"
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value "$($item.LocalizedDisplayName)::$($DeploymentType.Title.InnerText)"
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath ($DeploymentType.Installer.Contents.Content.Location).TrimEnd('\')
                    $GetAllPathsResult = Get-AllPaths -Path $obj.SourcePath -Cache $ShareCache
                    $ShareCache = $GetAllPathsResult[0]
                    $AllPaths = $GetAllPathsResult[1]
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                    $AllContent += $obj
                }
            }
            "Get-CMDriver" {
                $obj = New-Object PSObject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.CI_ID
                Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.LocalizedDisplayName
                Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath ($item.ContentSourcePath).TrimEnd('\')
                $GetAllPathsResult = Get-AllPaths -Path $obj.SourcePath -Cache $ShareCache
                $ShareCache = $GetAllPathsResult[0]
                $AllPaths = $GetAllPathsResult[1]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                $AllContent += $obj
            }
            default {
                $obj = New-Object PSObject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentType -Value ($Command -replace "Get-CM")
                Add-Member -InputObject $obj -MemberType NoteProperty -Name UniqueID -Value $item.PackageId
                Add-Member -InputObject $obj -MemberType NoteProperty -Name Name -Value $item.Name
                # OS images and boot iamges are absolute paths to files
                If ("OperatingSystemImage","BootImage" -contains $obj.ContentType) {
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath (Split-Path $item.PkgSourcePath).TrimEnd('\')
                }
                Else {
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SourcePath ($item.PkgSourcePath).TrimEnd('\')
                }
                $GetAllPathsResult = Get-AllPaths -Path $obj.SourcePath -Cache $ShareCache
                $ShareCache = $GetAllPathsResult[0]
                $AllPaths = $GetAllPathsResult[1]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name AllPaths -Value $AllPaths
                $AllContent += $obj
            }   
        }
    }
}
$ShareCache
$AllContent