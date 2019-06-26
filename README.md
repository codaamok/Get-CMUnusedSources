# Get-CMUnusedSources

## Table of contents

## Description

A PowerShell script that will tell you what folders are not used by System Center Configuration Manager content objects in a given path. This is useful if your storage is getting full and you need a way to identify what on disk is good to go.

The script returns an array of PSObjects where each PSObject has two properties:  `Folder` and `UsedBy`. The `UsedBy` prop can have one or more of the following values:

- A list all of the names for the content objects used by the folder.*
- `An intermediate folder (sub or parent folder)`: a sub or parent folders of a folder used by a content object.
- `Access denied`: a folder that the user running the script does not have read access to.
- `Not used`: a folder that is not used by any content objects.

\* for Applications, the naming convention will be `<Application name>::<Deployment Type name>`.

> **Note:** I refer to a "content object" as an object within ConfigMgr that has a source path associated with it i.e. an object with content. This includes:
> - Packages
> - Applications
> - Drivers
> - Driver packages
> - Operating System images
> - Operating System upgrade images
> - Boot images
> - Deployment packages

## Requirements

- ConfigMgr console installed
- PowerShell 5.1 or newer
- [PSWriteHTML](https://github.com/EvotecIT/PSWriteHTML) module installed (only if you specify `-HtmlReport` switch)

## Getting started

1. Download Get-CMUnusedSources.ps1
2. Check out the examples and read through the parameters available. If you're eager, then calling the script as is and only providing values for the mandatory parameters will get you going under the default conditions (_see below_).

## Default conditions

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\server\folder" -SiteCode "XYZ" -SiteServer "server.contoso.com"
```

Running the script without anything other than the mandatory parameters will do the following:

- Gather all content objects
- Show overall progress using Write-Progress
- No logging
- No PowerShell object export
- No HTML report
- Number of threads will be number of cores minus 1

## What can it do

This script can be run from your desktop and makes no changes. It's purely for reporting. It returns an exportable PowerShell object, because what good is a script that's for reporting where you can't actually do anything with the results? Finally, you can optional create a HTML report where you can then export to CSV/PDF/XSLX.

- The script returns the results as an array of PSObjects. I find it useful when a script that's used for reporting returns something that I can immediately do _something_ with.
- -SourcesLocation can be a UNC or local path. Do not worry about the MAX_PATH limit as the script prefixes what you give with `\\?\UNC\..` where the MAX_PATH limit is 32767 - [more info](https://docs.microsoft.com/en-us/windows/desktop/fileio/naming-a-file#maximum-path-length-limitation).
- You can filter the content object search by specifying one or more of the following:  `-Applications`,  `-Packages`,  `-Drivers`,  `-DriverPackages`,  `-OSImages`,  `-OSUpgradeImages`,  `-BootImages`,  `-DeploymentPackages`.
- Surpress the use of Write-Progress.
- Output a log file, enable log rotation, set a maximum log file size and how many rotated log files to keep.
- Export the array that's returned by the script to file. You can reimport it later using `Import-Clixml`.
- Export the result to HTML, and thanks to [PSWriteHTML](https://github.com/EvotecIT/PSWriteHTML), from there you can export to CSV/PDF/XSLX.
- The script uses runspaces so you can control how many threads are concurrently used.

## Examples

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\server\folder" -SiteCode "XYZ" -SiteServer "server.contoso.com" -Log -LogFileSize 2MB -ExportReturnObject -HtmlReport -Threads 2
```

It will gather all content objects relevant to site code `XYZ` and all folders under `\\server\folder`. A log file will be created in the same directory as the script and rolled over when it reaches 2MB, with no limit on number of rotated logs to keep. When finished, the object returned by the script will be exported and also the HTML report too. 2 threads will be used.

---

```powershell
PS C:\>
```

Another example here

---

```powershell
PS C:\>
```

Another example here

## Runtime stats

6 threads on an 8 core VM with 4GB RAM, x many content objects and x many folders:
_leave -htmlreport out_

## Process overview

The process begins by by gathering all (or selective) content objects within a hierarchy by site code using the ConfigMgr cmdlets. It also recursively gathers all folders under a given path.

The source path for each content object is manipulated to create every possible valid permutation. For example, say PackageABC1 has a source path `\\server\Applications$\7zip\x64`. The script will create an AllPaths property with a list like this:

```text
\\192.168.175.11\Applications$\7zip\x64
\\192.168.175.11\F$\Applications\7zip\x64
\\192.168.175.11\SomeOtherSharedFolder$\7zip\x64
\\server.contoso.com\Applications$\7zip\x64
\\server.contoso.com\F$\Applications\7zip\x64
\\server.contoso.com\SomeOtherSharedFolder$\7zip\x64
\\server\Applications$\7zip\x64
\\server\F$\Applications\7zip\x64
\\server\SomeOtherSharedFolder$\7zip\x64
F:\Applications\7zip\x64
````

In the above example, the script discovered the local path for the `Applicatons$` share was `F:\Applications` and that `SomeOtherSharedFolder$` was another share that also reoslves to the same local path. This path permutation enables the script to identify used folders that could use different absolute paths but resolve to the same folder.

Once all the folders and ConfigrMgr content objects have been gathered, it begins iterating through through each folder, and for each folder it iterates over all content objects to determine if said folder is used. This process builds an array of PSObjects and returns said array once complete.

### Example output

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\fileserver\Applications$" -SiteCode "XYZ" -SiteServer "server.contoso.com" -Applications
Starting
Gathering folders: \\fileserver\Applications$
Number of folders: 39
Gathering content objects: Application
Number of content objects: 16
Determinig unused folders, using 2 threads
Total number of content objects: 16
Total number of folders at \\fileserver\Applications$: 39
Total number of folders where access denied: 0
Total number of folders unused: 10
Total runtime: 00:00:07.2891577

PS C:\> $result | Select -First 10

Folder                                    UsedBy
------                                    ------
\\fileserver\Applications$                An intermediate folder (sub or parent folder)
\\fileserver\Applications$\.NET3.5        Not used
\\fileserver\Applications$\.NET3.5SP1     Not used
\\fileserver\Applications$\7zip           An intermediate folder (sub or parent folder)
\\fileserver\Applications$\7zip\uninstall 7-Zip 19.00::7-Zip 19.00 - Windows Installer (*.msi file)
\\fileserver\Applications$\7zip\x64       7-Zip 19.00::7-Zip 19.00 (x64 edition) - Windows Installer (*.msi file)
\\fileserver\Applications$\7zip\x86       Not used
\\fileserver\Applications$\chrome         An intermediate folder (sub or parent folder)
\\fileserver\Applications$\chrome\x64     Chrome 73.0.3683.103::Google Chrome x64
\\fileserver\Applications$\chrome\x86     Chrome 73.0.3683.103::Google Chrome x86
```

## Parameters

### -SourcesLocation (mandatory)

The path to the directory you store your ConfigMgr sources. Can be a UNC or local path. Must be a valid path that you have read access to.

### -SiteCode (mandatory)

The site code of the ConfigMgr site you wish to query for content objects.

### -SiteServer (mandatory)

The site server of the given ConfigMgr site code. The server must be reachable over a network.

### -Packages

Specify this switch to include Packages within the search to determine unused content on disk.

### -Applications

Specify this switch to include Applications within the search to determine unused content on disk.

### -Drivers

Specify this switch to include Drivers within the search to determine unused content on disk.

### -DriverPackages

Specify this switch to include DriverPackages within the search to determine unused content on disk.

### -OSImages

Specify this switch to include OSImages within the search to determine unused content on disk.

### -OSUpgradeImages

Specify this switch to include OSUpgradeImages within the search to determine unused content on disk.

### -BootImages

Specify this switch to include BootImages within the search to determine unused content on disk.

### -DeploymentPackages

Specify this switch to include DeploymentPackages within the search to determine unused content on disk.

### -AltFolderSearch

Specify this if you suspect there are issue with the default mechanism of gathering folders, which is:

```powershell
Get-ChildItem -LiteralPath "\\?\UNC\server\share\folder" -Directory -Recurse | Select-Object -ExpandProperty FullName
```

### -NoProgress

Specify this to disable use of Write-Progress.

### -Log

Specify this to enable logging. The log file(s) will be saved to the same directory as this script with a name of `<scriptname>_<datetime>.log`. Rolled log files will follow a naming convention of `<filename>_1.lo_` where the int increases for each rotation.

### -LogFileSize

Set the maximum size you want for each rolled over log file. This is only applicable if NumOfRotatedLogs is greater than 0. Default value is 5MB. The unit of measurement is bytes however you can specify units such as KB, MB etc.

### -NumOfRotatedLogs

Set the maximum number of log files you wish to keep. Default value is 5MB. Specify 0 for unlimited.

### -ExportReturnObject

Specify this option if you wish to export the PowerShell result object to an XML file. The XML file be saved to the same directory as this script with a name of `<scriptname>_<datetime>_result.xml`. It can easily be reimported using Import-Clixml cmdlet.

### -ExportCMContentObjects

Specify this option if you wish to export all ConfigMgr content objects to an XML file. The XML file be saved to the same directory as this script with a name of `<scriptname>_<datetime>_cmobjects.xml`. It can easily be reimported using Import-Clixml cmdlet.

### -HtmlReport

Specify this option to enable the generation for a HTML report of the result. Doing this will force you to have the PSWriteHtml module installed. For more information on [PSWriteHTML](https://github.com/EvotecIT/PSWriteHTML). The HTML file will be saved to the same directory as this script with a name of `<scriptname>_<datetime>.html`.

### -Threads

Set the number of threads you wish to use for concurrent processing of this script. Default value is number of processes from env var NUMBER_OF_PROCESSORS minus 1.

## Author

Adam Cook

Twitter: [@codaamok](https://twitter.com/codaamok)

Website: https://www.cookadam.co.uk

## License

Please read the LICENSE file provided.

## Acknowledgements

Big thanks to folks in [Windows Admins slack](https://slofile.com/slack/winadmins):

- Cody Mathis ([@codymathis123](https://github.com/CodyMathis123))
- Chris Kibble ([@ChrisKibble](https://github.com/ChrisKibble))
- Chris Dent ([@idented-automation](https://github.com/indented-automation))
- Kevin Crouch ([@PsychoData](https://github.com/PsychoData))
- Patrick (the guy who wrote [MakeMeAdmin](https://makemeadmin.com/))
