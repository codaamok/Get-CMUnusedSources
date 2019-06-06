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
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Server -ErrorAction Stop
        }
        # Set the current location to be the site code.
        Set-Location "$($SiteCode):\" -ErrorAction Stop

        # Verify given sitecode
        If((Get-CMSite -SiteCode $SiteCode).SiteCode -ne $SiteCode) { throw }

    } catch {
        If((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -ne $null) {
            Set-Location $Path
            Remove-PSDrive -Name $SiteCode -Force
        }
        Throw "Failed to create New-PSDrive with site code `"$($SiteCode)`" and server `"$($Server)`""
    }

}

Set-CMDrive -SiteCode "ACC" -Server "S2CCM" -Path "C:\Users\Adam"