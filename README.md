# Get-CMUnusedSources

## Table of contents

1. [Description](#description)
2. [Requirements](#requirements)
3. [Getting started](#getting-started)
4. [Default conditions](#default-conditions)
5. [What can it do](#what-can-it-do)
6. [Examples](#examples)
7. [Runtime stats](#runtime-stats)
8. [Process overview](#process-overview)
9. [Validating the results](#validating-the-results)
10. [The Excel report explained](#the-excel-report-explained)
11. [The log file explained](#the-log-file-explained)
12. [Parameters](#parameters)
13. [Author](#author)
14. [License](#license)
15. [Acknowledgements](#acknowledgements)

## Description

A PowerShell script that will tell you what folders are not used by System Center Configuration Manager in a given path. This is useful if your storage is getting full and you need a way to identify what on disk is good to go.

Watch [this video](https://www.youtube.com/watch?v=YGwQIUhYJsY) of me demoing it at WMUG.

The script returns an array of PSObjects with two properties: `Folder` and `UsedBy`.

The `UsedBy` property can have one or more of the following values:

- A list of names for the content objects used by the folder.*
- `An intermediate folder (sub or parent folder)`: a sub or parent folders of a folder used by a content object.
- `Access denied` or `Access denied (elevation required`): a folder that the user running the script does not have read access to.
- `Not used`: a folder that is not used by any content objects.

\* for Applications, the naming convention will be `<Application name>::<Deployment Type name>`.

For example:

```powershell
PS C:\> $result | Select -First 10

Folder                                    UsedBy
------                                    ------
\\fileserver\Applications$                An intermediate folder (sub or parent folder)
\\fileserver\Applications$\.NET3.5        Not used
\\fileserver\Applications$\.NET3.5SP1     Not used
\\fileserver\Applications$\7zip           An intermediate folder (sub or parent folder)
\\fileserver\Applications$\7zip\uninstall 7-Zip 19.00::7-Zip 19.00 - Windows Installer (*.msi file)
\\fileserver\Applications$\7zip\x64       7-Zip 19.00::7-Zip 19.00 (x64 edition) - Windows Installer (*.msi file)
\\fileserver\Applications$\7zip\x86       7-Zip 19.00::7-Zip 19.00 - Windows Installer (*.msi file)
\\fileserver\Applications$\chrome         An intermediate folder (sub or parent folder)
\\fileserver\Applications$\chrome\x64     Chrome 73.0.3683.103::Google Chrome x64
\\fileserver\Applications$\chrome\x86     Chrome 73.0.3683.103::Google Chrome x86
```

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
- [ImportExcel](https://github.com/dfinke/ImportExcel) module installed (only if you specify `-ExcelReport` switch). See an example of the Excel report [here](https://www.cookadam.co.uk/Get-CMUnusedSources.ps1.xlsx).

## Getting started

1. Download `Get-CMUnusedSources.ps1`
2. Check out the examples below and read through the parameters available. If you're eager, then calling the script as is and only providing values for the mandatory parameters will get you going under the default conditions (_see below_).

## Default conditions

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\server\folder" -SiteServer "server.contoso.com"
```

Running the script without anything other than the mandatory parameters will do the following:

- Gather all content objects under the site code of site server given in -SiteServer
- Show overall progress using Write-Progress
- No logging
- No PowerShell object export
- No Excel report
- Number of threads will be number from environment variable `NUMBER_OF_PROCESSORS`

## What can it do

- You can execute this script remotely from a site server. It makes zero changes to your ConfigMgr site. It's purely for reporting.
- The returned object by the script is an array of PSObjects.
- `-SourcesLocation` can be a UNC or local path and works around the 260 MAX_PATH limit. Use `-ExcludeFolders` parameter which takes an array of absolute paths of folders you wish to exclude under `-SourcesLocation`.
- `-Threads` to control how many threads are used for concurrent processing.
- Optionally exports PowerShell objects to file of either all your ConfigMgr content objects or the result. You can later reimport these using `Import-Clixml`.
- Optionally create an Excel report.
- Optionally filter the ConfigMgr content object search by specifying one or more of the following:  `-Applications`,  `-Packages`,  `-Drivers`,  `-DriverPackages`,  `-OSImages`,  `-OSUpgradeImages`,  `-BootImages`,  `-DeploymentPackages`.
- Optionally create a log file.
- Optionally produce an Excel report, with thanks to [ImportExcel](https://github.com/dfinke/ImportExcel). See an example of the Excel report [here](https://www.cookadam.co.uk/Get-CMUnusedSources.ps1.xlsx).

## Examples

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\server\folder" -SiteServer "server.contoso.com" -Log -ExportReturnObject -ExcelReport -ExcludeFolders "\\server\folder\somechildfolder1", "\\server\folder\somechildfolder2" -Threads 2
```

- Gather all content objects relevant to site code `XYZ`.
- Gather all folders under `\\server\folder`.
- Exclude gathering (and therefore later processing) folders under `\\server\folder\somechildfolder1` and `\\server\folder\somechildfolder2`.
- A log file will be created in the same directory as the script and rolled over when it reaches 2MB, with no limit on number of rotated logs to keep.
- When finished, the object returned by the script will be exported and also the Excel report too.
- 2 threads will be used.
- Returns the result PowerShell object to variable `$result`.

---

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "F:\some\folder" -SiteServer "server.contoso.com" -Log -NoProgress -ExportReturnObject -ExportCMContentObjects -Packages -Applications -OSImages -OSUpgradeImages -ExcelReport
```

- Gather all content objects relevant to site code `XYZ`.
- Gather all folders under `F:\some\folder`.
- A log file will be created in the same directory as the script and rolled over when it reaches 2MB, with no limit on number of rotated logs to keep.
- Suppress PowerShell progress (`Write-Progress`).
- Exports the result PowerShell object to file saved in the same directory as the script.
- Exports all searched ConfigMgr content objects to file saved in the same directory as the script.
- Gathers only Packages, Applications, Operating System images and Operating System upgrade images content objects.
- Produces an Excel report saved in the same directory as the script. See an example of the Excel report [here](https://www.cookadam.co.uk/Get-CMUnusedSources.ps1.xlsx).
- Will use as many threads as the value in environment variable `NUMBER_OF_PROCESSORS` because that's the default value of `-Threads`.
- Returns the result PowerShell object to variable `$result`

## Runtime stats

The below stats are an average of 3 runs using the following options:

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\server\folder" -SiteServer "server.contoso.com"
```

**Folders:** 2633 - **Content objects:** 132 - **CPUs:** 2 - **RAM:** 8GB - **Runtime:** 4 minutes 20 seconds

... more come

## Process overview

The process begins by gathering all folders under `-SourcesLocation`. Once all folder are gathered, all (or selective) ConfigMgr content objects are gathered within a hierarchy, filtered by site code, using the ConfigMgr cmdlets.

> **Note:** It's OK if you see "Access denied" exceptions printed to console, the script will handle these and report accordingly.

The source path for each content object is manipulated to create every possible valid permutation. For example, say package P01000CD has a source path `\\server\Applications$\7zip\x64`. The script will create an AllPaths property with a list like this:

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

In the above example, the script discovered the local path for the `Applications$` share was `F:\Applications` and that `SomeOtherSharedFolder$` was another share that also resolves to the same local path. This path permutation enables the script to identify used folders that could use different paths but resolve to the same location.

Once all the folders and ConfigMgr content objects have been gathered, it begins iterating through through each folder, and for each folder it iterates over all content objects to determine if said folder is used by any content objects. 

This process builds an array of PSObjects with the two properties `Folder` and `UsedBy`.

The `UsedBy` property can have one or more of the following values:

- A list of names for the content objects used by the folder.*
- `An intermediate folder (sub or parent folder)`: a sub or parent folders of a folder used by a content object.
- `Access denied` or `Access denied (elevation required)`: a folder that the user running the script does not have read access to.
- `Not used`: a folder that is not used by any content objects.

\* for Applications, the naming convention will be `<Application name>::<Deployment Type name>`.

### Example output

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\fileserver\Applications$" -SiteServer "server.contoso.com" -Applications
[ 00:18:14 | 00:00:00 ] - Starting
[ 00:18:14 | 00:00:00 ] - Gathering folders: \\fileserver\Applications$
[ 00:18:14 | 00:00:00 ]   - Done, number of gathered folders: 216
[ 00:18:14 | 00:00:00 ] - Gathering content objects: Application
[ 00:18:19 | 00:00:05 ]   - Done, number of gathered content objects: 74
[ 00:18:19 | 00:00:05 ] - Determining unused folders, using 4 threads
[ 00:18:19 | 00:00:05 ]   - Adding jobs to queue
[ 00:18:23 | 00:00:09 ]       - Done, waiting for jobs to complete
[ 00:18:31 | 00:00:17 ]   - Done determining unused folders
[ 00:18:32 | 00:00:17 ] - Calculating used disk space by unused folders
[ 00:18:32 | 00:00:17 ]   - Done calculating used disk space by unused folders
[ 00:18:32 | 00:00:17 ] - ---------------------------------------------------------------------------
[ 00:18:32 | 00:00:17 ] - Folders in \\fileserver\Applications$: 216
[ 00:18:32 | 00:00:17 ] - Folders where access denied: 2
[ 00:18:32 | 00:00:17 ] - Folders unused: 64
[ 00:18:32 | 00:00:17 ] - Potential disk space savings in "\\fileserver\Applications$": 1661.43 MB
[ 00:18:32 | 00:00:17 ] - Content objects processed: Application
[ 00:18:32 | 00:00:17 ] - Content objects: 74
[ 00:18:32 | 00:00:17 ] - Runtime: 00:00:17.7590537
[ 00:18:32 | 00:00:17 ] - ---------------------------------------------------------------------------
[ 00:18:32 | 00:00:17 ] - Finished

PS C:\> $result | Select -First 10

Folder                                    UsedBy
------                                    ------
\\fileserver\Applications$                An intermediate folder (sub or parent folder)
\\fileserver\Applications$\.NET3.5        Not used
\\fileserver\Applications$\.NET3.5SP1     Not used
\\fileserver\Applications$\7zip           An intermediate folder (sub or parent folder)
\\fileserver\Applications$\7zip\uninstall 7-Zip 19.00::7-Zip 19.00 - Windows Installer (*.msi file)
\\fileserver\Applications$\7zip\x64       7-Zip 19.00::7-Zip 19.00 (x64 edition) - Windows Installer (*.msi file)
\\fileserver\Applications$\7zip\x86       7-Zip 19.00::7-Zip 19.00 - Windows Installer (*.msi file)
\\fileserver\Applications$\chrome         An intermediate folder (sub or parent folder)
\\fileserver\Applications$\chrome\x64     Chrome 73.0.3683.103::Google Chrome x64
\\fileserver\Applications$\chrome\x86     Chrome 73.0.3683.103::Google Chrome x86
```

## Validating the results

You could verify if a folder structure is used or not by checking out the results of the script. The Excel report has a tab "Content objects" where you can verify if part or all of the path you're curious about is used.

You can easily achieve this verification with PowerShell:

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\fileserver\Applications$" -SiteServer "server.contoso.com" -Applications
[ 00:18:14 | 00:00:00 ] - Starting
[ 00:18:14 | 00:00:00 ] - Gathering folders: \\fileserver\Applications$
[ 00:18:14 | 00:00:00 ]   - Done, number of gathered folders: 216
[ 00:18:14 | 00:00:00 ] - Gathering content objects: Application
[ 00:18:19 | 00:00:05 ]   - Done, number of gathered content objects: 74
[ 00:18:19 | 00:00:05 ] - Determining unused folders, using 4 threads
[ 00:18:19 | 00:00:05 ]   - Adding jobs to queue
[ 00:18:23 | 00:00:09 ]       - Done, waiting for jobs to complete
[ 00:18:31 | 00:00:17 ]   - Done determining unused folders
[ 00:18:32 | 00:00:17 ] - Calculating used disk space by unused folders
[ 00:18:32 | 00:00:17 ]   - Done calculating used disk space by unused folders
[ 00:18:32 | 00:00:17 ] - ---------------------------------------------------------------------------
[ 00:18:32 | 00:00:17 ] - Folders in \\fileserver\Applications$: 216
[ 00:18:32 | 00:00:17 ] - Folders where access denied: 2
[ 00:18:32 | 00:00:17 ] - Folders unused: 64
[ 00:18:32 | 00:00:17 ] - Potential disk space savings in "\\fileserver\Applications$": 1661.43 MB
[ 00:18:32 | 00:00:17 ] - Content objects processed: Application
[ 00:18:32 | 00:00:17 ] - Content objects: 74
[ 00:18:32 | 00:00:17 ] - Runtime: 00:00:17.7590537
[ 00:18:32 | 00:00:17 ] - ---------------------------------------------------------------------------
[ 00:18:32 | 00:00:17 ] - Finished

PS C:\> $result | Where-Object { $_.Folder -like "\\fileserver\Applications$*" }

Folder                                UsedBy
------                                ------
\\fileserver\Applications$\.NET3.5    Not used
\\fileserver\Applications$\.NET3.5SP1 Not used
\\fileserver\Applications$\Office     Not used
\\fileserver\Applications$\Office\x64 Not used
\\fileserver\Applications$\Office\x86 Not used
...
```

The `-ExportCMContentObjects` is also a useful switch for this task because it produces an XML export of the result produced by `Get-CMContent`. `Get-CMContent` is a function within the script that retrieves content objects using the Configuration Manager PowerShell cmdlets and selecting only key properties.

For all the content objects `Get-CMContent` gathers, the following properties are selected:

- `ContentType`
- `UniqueID`
- `Name`
- `SourcePath`
- `SourcePathFlag`
- `AllPaths`
- `SizeMB`

The `SourcePathFlag` property is an enum which can have the following fours values:

- `0` = `[FileSystemAccessState]::ERROR_SUCCESS`
- `3` = `[FileSystemAccessState]::ERROR_PATH_NOT_FOUND`
- `5` = `[FileSystemAccessState]::ERROR_ACCESS_DENIED`
- `740` = `[FileSystemAccessState]::ERROR_ELEVATION_REQUIRED`

Below is an example of how to generate the result into XML and import it in to a different variable. From there you can do any filter you wish using `.Where()` or `Where-Object { .. }`.

```powershell
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\fileserver\Applications$" -SiteServer "server.contoso.com" -Applications -ExportCMContentObjects
...
PS C:\> $cmcontentobjs = Import-Clixml -Path ".\Get-CMUnusedSources.ps1_2019-07-07_16-48-52_cmobjects.xml"
PS C:\> $cmcontentobjs | Select -First 2

ContentType    : Application
UniqueID       : DeploymentType_0ab33a06-96ee-441a-83d8-3d3b6d0be224
Name           : Chrome 73.0.3683.103::Google Chrome x86
SourcePath     : \\fileserver.contoso.com\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x86\
SourcePathFlag : 0
AllPaths       : {\\192.168.175.11\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x86,
                 \\fileserver\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x86,
                 \\fileserver.contoso.com\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x86,
                 \\fileserver.contoso.com\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x86...}
SizeMB         : 56.48


ContentType    : Application
UniqueID       : DeploymentType_f77958c3-2d93-4ac7-a81d-02e83aa69b83
Name           : Chrome 73.0.3683.103::Google Chrome x64
SourcePath     : \\fileserver.contoso.com\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x64\
SourcePathFlag : 0
AllPaths       : {\\fileserver\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x64,
                 \\192.168.175.11\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x64,
                 \\fileserver.contoso.com\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x64,
                 \\fileserver.contoso.com\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x64...}
SizeMB         : 57.36
```

The `AllPaths` property is a hashtable.

```powershell
PS C:\> $cmcontentobjs | Select-Object -ExpandProperty AllPaths -First 2 | Select-Object -ExpandProperty Keys

\\192.168.175.11\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver.contoso.com\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver.contoso.com\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver.contoso.com\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x86
\\192.168.175.11\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x86
\\192.168.175.11\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver.contoso.com\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x86
\\192.168.175.11\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x86
F:\Applications\chrome\chrome 73.0.3683.103\Google Chrome x86
\\fileserver\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x64
\\192.168.175.11\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver.contoso.com\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver.contoso.com\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x64
\\192.168.175.11\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x64
F:\Applications\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver\Sources$\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver.contoso.com\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x64
\\192.168.175.11\Applications1992\chrome\chrome 73.0.3683.103\Google Chrome x64
\\fileserver.contoso.com\F$\Applications\chrome\chrome 73.0.3683.103\Google Chrome x64
\\192.168.175.11\Applications$\chrome\chrome 73.0.3683.103\Google Chrome x64
```

## The Excel report explained

See an example of the Excel report [here](https://www.cookadam.co.uk/Get-CMUnusedSources.ps1.xlsx).

You'll see five tabs: `Result`, `Summary`, `Not used folders`, `Content objects` and `Invalid paths`.

### Result

All of the folders under `-SourcesLocation` and their UsedBy status (same as the PSObject returned by the script).

### Summary

A list of folders that were determined not used under the given path by the searched content objects. It does not include child folders, only "unique root folders", so this produces an accurate measurement of capacity used.

For example, if `\\server\share\7zip`, `\\server\share\7zip\x64` and `\\server\share\7zip\x86` were all marked as "Not used" then this view will only list `\\server\share\7zip`. 

This ensures the "Totals" column produces an accurate measure of total capacity, file count or directory account for all folders marked as "Not used".

### All not used folders

All folders that were determined not used under the given path by the searched content objects.

### Invalid paths

Content objects that have a source path which are not accessible from the computer that ran this script.

A common result you'll see here, if you run the script remote from a site server, is things like USMT packages that are traditionally stored as a Package with a local path rather than UNC path.

### Content objects

All searched ConfigMgr content objects. For example, if you specified `-Drivers` and `-DriverPackages` then it would only show you Driver and DriverPackage content object types, because that's all that was gathered.

For each content object, you'll see these properties:

- `ContentType`
- `UniqueID`
- `Name`
- `SourcePath`
- `SourcePathFlag`
- `AllPaths`
- `SizeMB`

## The log file explained

You'll need CMTrace to read it.

The log will display everything you see printed to console as well as gathered content object properties and the result.

After gathering each content object, it will log all of those content objects and their properties like so: 

> ContentType - UniqueId - Name - SourcePath - SourcePathFlag - AllPaths - SizeMB

Example: 

> Application - DeploymentType_1cd4198d-bb1f-41a8-ad3e-a81f4d8f1d44 - PuTTY 0.71::PuTTY x64 - \\\\fileserver.contoso.com\\Applications$\\putty\\putty 0.71\\PuTTY x64\\ - 0 - \\\\fileserver\\Applications$\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver.contoso.com\\F$\\Applications\\putty\\putty 0.71\\PuTTY x64,F:\\Applications\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver.contoso.com\\Applications$\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver\\Applications1992\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver.contoso.com\\Sources$\\putty\\putty 0.71\\PuTTY x64,\\\\192.168.175.11\\Sources$\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver\\Sources$\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver\\F$\\Applications\\putty\\putty 0.71\\PuTTY x64,\\\\192.168.175.11\\Applications$\\putty\\putty 0.71\\PuTTY x64,\\\\192.168.175.11\\Applications1992\\putty\\putty 0.71\\PuTTY x64,\\\\192.168.175.11\\F$\\Applications\\putty\\putty 0.71\\PuTTY x64,\\\\fileserver.contoso.com\\Applications1992\\putty\\putty 0.71\\PuTTY x64 - 0.73

The result of the script is also written to log like so:

> Folder: UsedBy

Example:

```
...
\\fileserver\Applications$\python3: An intermediate folder (sub or parent folder)
\\fileserver\Applications$\python3\python3 3.7.3: An intermediate folder (sub or parent folder)
\\fileserver\Applications$\python3\python3 3.7.3\Python x64: Python 3.7.3::Python x64
\\fileserver\Applications$\python3\python3 3.7.3\Python x86: Python 3.7.3::Python x86
\\fileserver\Applications$\treesizefree: An intermediate folder (sub or parent folder)
\\fileserver\Applications$\treesizefree\treesizefree 4.3.1: An intermediate folder (sub or parent folder)
\\fileserver\Applications$\treesizefree\treesizefree 4.3.1\TreeSize Free: TreeSize Free 4.3.1::TreeSize Free
\\fileserver\Applications$\vlc: An intermediate folder (sub or parent folder)
\\fileserver\Applications$\vlc\vlc 3.06: An intermediate folder (sub or parent folder)
\\fileserver\Applications$\vlc\vlc 3.06\VLC x64: VLC Media Player 3.06::VLC x64
\\fileserver\Applications$\vlc\vlc 3.06\VLC x86: VLC Media Player 3.06::VLC x86
...
```

You may spot some highlighted yellow warnings or red errors so I'll explain some of them here.

> Unable to interpret path "x"

Occurs during the gathering content objects stage and trying to build the AllPaths property.

One of your content objects in ConfigMgr has a funky source path that I couldn't determine. If you get this, please share with me "x"!

This would only be problematic for the content object(s) that experience this.

> Server "x" is unreachable

Occurs during the gathering content objects stage and trying to build the AllPaths property.

The server "x" that hosts the shared folder is unreachable; it failed a ping test.

This would only be problematic for the content object(s) that experience this.

> Could not query Win32_Share on "x" (y)

Occurs during the gathering content objects stage and trying to build the AllPaths property.

Could not query the Win32_Share WMI class on the server "x" that hosts the shared folder. "y" is the exception message.

This would only be problematic for the content object(s) that have source path associated with this server and any of its shared folders.

> Could not resolve share "x" on "y", either because it does not exist or could not query Win32_Share on server

Occurs during the gathering content objects stage and trying to build the AllPaths property.

Could not determine the local path of shared folder "x" because that information could not be retrieved from server "y", either because the share does not exist or could not query the Win32_Share class on "y".

This would only be problematic for the content object(s) that experience this.

> Couldn't determine path type for "x" so might have problems accessing folders that breach MAX_PATH limit, quitting...

Occurs during the gathering folders stage so we can prefix `-SourcesLocation` with `\\?\..` to workaround the 260 MAX_PATH limit.

This means that `-SourcesLocation` was in such bananas format that I could not determine it. Please share "x" with me if you experience this!

This is a terminating error. It's important for the script to be able to successfully gather all folders from `-SourcesLocation`.

> Consider using -AltFolderSearch, quiting...

Occurs during the gathering folders stage.

Something went wrong trying to get all folders under `-SourcesLocation`.

This is a terminating error. It's important for the script to be able to successfully gather all folders from `-SourcesLocation`.

> Couldn't reset "x"

Occurs during the gathering folders stage so we can prefix `-SourcesLocation` with `\\?\..` to workaround the 260 MAX_PATH limit.

After the script has finished gathering folders it tries to remove the `\\?\..` prefix. If you experience this when this removal of the prefix failed for some reason. Please share "x" with me if you experience this!

This should not impact the results.

> Won't be able to determine unused folders with given local path while running remotely from site server, quitting

Occurs before gathering any content objects or folders.

You have given local path for `-SourcesLocation`. We need to ensure we don't produce false positives where a similar folder structure exists on the remote machine and site server.

This is a terminating error.

> Unable to import ImportExcel module: "x"

Occurs before gathering any content objects or folders.

You have specified the switch to produce Excel report but do not have the [ImportExcel](https://github.com/dfinke/ImportExcel) module installed.

This is a terminating error.

> Failed to export PowerShell object: "x"

Occurs after main execution and just before closing.

You have specified the switch to export either all ConfigMgr content objects or the result object, and there was a problem doing so.

The script will attempt to continue and close normally.

> Failed to create Excel report: "x"

Occurs after main execution and just before closing.

You have specified the switch to produce a Excel report and there was a problem doing so.

The script will attempt to continue and close normally.

## Parameters

### -SourcesLocation (mandatory)

The path to the directory you store your ConfigMgr sources. Can be a UNC or local path. Must be a valid path that you have read access to.

### -SiteServer (mandatory)

The site server of the given ConfigMgr site code. The server must be reachable over a network.

### -SiteCode

The site code of the ConfigMgr site you wish to query for content objects.

### -ExcludeFolders

An array of folders that you want to exclude the script from checking, which should be absolute paths under the path given for -SourcesLocation.

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

Specify this to enable logging. The log file(s) will be saved to the same directory as this script with a name of `<scriptname>_<datetime>.log`. Rolled log files will follow a naming convention of `<filename>_1.lo_` where the int increases for each rotation. Each maximum log file is 2MB.

### -ExportReturnObject

Specify this option if you wish to export the PowerShell result object to an XML file. The XML file be saved to the same directory as this script with a name of `<scriptname>_<datetime>_result.xml`. It can easily be reimported using `Import-Clixml` cmdlet.

### -ExportCMContentObjects

Specify this option if you wish to export all ConfigMgr content objects to an XML file. The XML file be saved to the same directory as this script with a name of `<scriptname>_<datetime>_cmobjects.xml`. It can easily be reimported using `Import-Clixml` cmdlet.

### -ExcelReport

Specify this option to enable the generation for an Excel report of the result. Doing this will force you to have the ImportExcel module installed. For more information on ImportExcel: https://github.com/dfinke/ImportExcel. The .xlsx file will be saved to the same directory as this script with a name of <scriptname>_<datetime>.xlsx.

### -Threads

Set the number of threads you wish to use for concurrent processing of this script. Default value is number of processes from environment variable `NUMBER_OF_PROCESSORS`.

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
- Chris Dent ([@indented-automation](https://github.com/indented-automation))
- Kevin Crouch ([@PsychoData](https://github.com/PsychoData))
- Patrick Seymour ([@pseymour](https://github.com/pseymour))
