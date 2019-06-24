# Get-CMUnusedSources

## Table of contents?

## Short Description

A PowerShell script that will tell you what folders are not used by System Center Configuration Manager in a given path. This is useful if your storage is getting full and you need a way to identify what on disk is good to go.

This script can be run from your desktop and makes no changes. It's purely for reporting. It returns an exportable PowerShell object or a HTML report where you can then export to CSV/PDF/XSLX.

## Requirements

- ConfigMgr console installed
- PowerShell 5.1
- [PSWriteHTML](https://github.com/EvotecIT/PSWriteHTML) module installed (only if you specify `-HtmlReport` switch)

## Getting started

1. Download Get-CMUnusedSources.ps1
2. Before executing, ensure at the very least you see the examples.

## Examples

The below will gather all content

```
PS C:\> $result = .\Get-CMUnusedSources.ps1 -SourcesLocation "\\server\folder" -SiteCode "XYZ" -SiteServer "server.contoso.com" -Log -LogFileSize 2MB -ObjectExport -HtmlReport -Threads 2
```

## Description

> **Note:** I refer to a "content object" as an object within ConfigMgr that has a source path associated with it i.e. an object with content.

It works by gathering all (or selective) content objects within a hierarchy by site code by using the ConfigMgr cmdlets. It also gathers all folders under a given path, by default it uses the `\\?\UNC` convention to avoid the 255 path limit, where instead the limit is 32,767 characters - [more info](https://docs.microsoft.com/en-us/windows/desktop/fileio/naming-a-file#maximum-path-length-limitation).

The source path for each content object is observed and manipulated to create every possible valid permutation. For example, say PackageABC1 has a source path `\\server\Applications$\7zip\x64`. The script will create an AllPaths property with a list like this:

```
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

In the above example, the script discovered the local path for the share was F:\Applications and that SomeOtherSharedFolder$ was another share name for the same directory. This path permutation enables the script to identify used folders where content objects could use different (but yet, the same) paths.

Once all the information is gathered, it uses runspaces for multithreaded processing to iterate through each folder, and for each folder iterate over all content objects to determine if said folder is used or unused.

The script returns array of PSObjects with two properties:

1. Folder

The full name of the directory

2. UsedBy

Three possible values: the content object(s) that use the folder, "Access denied" and/or "Not used"

Other outputs are available to you, such as `-HtmlReport` where it creates a depedency on the [PSWriteHtml](https://github.com/EvotecIT/PSWriteHTML) module. From there, you can then export to CSV/PDF/XSLX. You can also use `-ObjectExport` to dump the PowerShell object that the script returns to an XML file of so you can later reimport.

## Runtime stats

## Author

## License

## Acknowledgements
