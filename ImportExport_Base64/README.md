# Code Protection : Base64 encode and decode PowerShell scripts.
 
If you split your Powershell scripts in 1 main PS1 script file and several "resource" (or library) PS1 files, if you want to protect a little bit the code in your "resource" files, this script is for you !
 
This Powershell script contains 2 functions:
 
- Export-PSResources : This function allows you to encode "resource" files in Base64 and save them in PSL (like Powershell Script Library) files.
- Import-PSResources : This function allows you to decode and then dot source content of PSL files in your main script.