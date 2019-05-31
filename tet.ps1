Function Get-AllFolders {
    Param(
        [string]$dirName
    )

    Write-Host "Working on $dirName"

    [System.Collections.ArrayList]$FolderList = @((Get-ChildItem -Path $dirName -Directory).FullName | Where-Object { $_ -ne $null })

    ForEach($Folder in $FolderList) {
        $FolderList += Get-AllFolders -dirName $Folder
    }

    return $FolderList
}