# IIS Log Analyzer v1.0

This is a graphical tool, developped in Powershell and using Windows Forms, to export Microsoft IIS W3C formated logs.
This tool filter logs by exporting the list of users and/or client IP connections on IIS Web or FTP sites by date.

It is possible to analyse local and remote computers which have Microsoft Internet Information Services installed.
By default the script is looking for default log files location, but it is possible to specify other IIS log folders.

This script was tested on Powershell v5 and v6.

Here are files included in this Gist :

FileName | Desciption
-------- | ----------
IISLogsAnalyzer.ps1 | Main script file
\_00\_Objects.ps1 | "Library" of custom Object Classes used in the main script
\_01\_Resources.ps1 | "Library" of resources (arrays, strings, etc.) used in the main script
\_02\_SharedFunctions.ps1 | "Library" of custom functions used in the main script
\_03\_GUI.ps1 | "Library" of Windows Forms Graphical Interface related code and functions used by the main script

It is possible to protect and encode the _\_nn\_Library.ps1_ files in PSL base64 files by using [ImpExp_B64.ps1](https://gist.github.com/B4tchM4n/74d5088d1e1c52c39724134c74d6b37b).