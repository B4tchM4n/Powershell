###################################################################
####                                                           ####
####        ####   #####   ###   ####      #   #  #####        ####
####        #   #  #      #   #  #   #     ## ##  #            ####
####        ####   ###    #####  #   #     # # #  ###          ####
####        #   #  #      #   #  #   #     #   #  #            ####
####        #   #  #####  #   #  ####      #   #  #####        ####
####                                                           ####
###################################################################
####                    by Fabian Caignard                     ####
###################################################################
####     IISWebStats.ps1 & RemIISWStats scripts User Guide     ####
###################################################################

1. Script Description:
----------------------

This script let you read all IIS Web (or FTP) logs and join them in
1 common file.
It also remove double entries, based on same Date, User, Client IP
and Browsed URLs.

The script save its results in a text output file. This file is
saved in the script directory.
It is possible to provide your own filename and location.

A log path can be provided to the script to force it to get logs
files from a specific location.
The provided log path can be a directory or a file.

If no log path is passed to the script, then it will search for
default IIS logs locations (C:\Inetpub\logs\LogFiles\ or
C:\Windows\system32\LogFiles\)
By default, when working on default IIS logs locations, the script
is looking for Web logs (W3SVC* directories). It is possible to
look for FTP logs (FTPSVC* directories).

Sometimes, it could occur that there are a lot of log files.
In this case the script could need to run for a long time.
If you need to limit only the analysis on a small period of time,
it possible to limit the scope of the analysis of 1 to 12 last
months of logs.

By default, the script is only analyze files with .log extension.
It is possible to force script to analyze a text file with an other
extension by passing it in parameters (see LogPath).

This script was tested on Powershell 2.0 to 5.1 : it is possible
that the script does not work properly on previous vervion of
Powershell.

Specific features and process for RemIISWStats remote script:
The script check first if computer passed in Server parameter is
reachable on the network, then, if Username and Password were
provided, it tries to check authentication.
If authentication failed then script exits.
After these checks, the script tries to connect on remote server
C drive (or drive/share root if LogPath provided).
If passed credentials (User and Password) are not permitted to
connect,or credentials are needed but no credentials were provided,
the script will ask if you want to provide other credentials.
If new connection try failed, then script exits.


2. How to run this script:
--------------------------

It is recommended to copy the script in a temporary directory on
the server on which you want to analyze logs.

	1. Open Micorsoft Powershell Console
	2. Browse to the script location
	3. type ".\IISWebStats.ps1" (without quotes)
	4. Press Enter key or check Syntax chapter.

	
3. Srcipts syntax:
-----------------

	Use IISWebStats script to run localy on IIS server:

		.\IISWebStats.ps1 [[-LogType] <W3|FTP>] [[-SinceLast] <1-12>] [[-OutFile] <string>]

		.\IISWebStats.ps1 [-LogPath] <string> [[-SinceLast] <1-12>] [[-OutFile] <string>]


	Use RemIISWStats script to run localy on any computer and remote connect to IIS Server:

		.\RemIISWStats.ps1 [[-Server] <string>] [[-Username] <string> [-Password] <string>] [[-LogType] <W3|FTP>] [[-SinceLast] <1-12>] [[-OutFile] <string>]

		.\RemIISWStats.ps1 [[-Server] <string>] [[-Username] <string> [-Password] <string>] [-LogPath] <string> [[-SinceLast] <1-12>] [[-OutFile] <string>]

	
	Parameters:
	
	Server		(String)	(Mandatory) Hostname of the remote IIS Server. This parameter can
							contain only 1 computer name and does not accept wildcards.
	Username	(String)	Username for remote connection authentication. It is recommended to
							provide username including domain name (Doamin\User or User@Domain).
	Password	(String)	(Mandatory when Username parameter is passed) Password corresponding
							to passed username.
	LogPath		(String)	(Mandatory) Path to a directory containing logs files or path to a
							log file.
							This parameter cannot contain wildcards and shoud contain a full
							path.
	LogType		(String)	Determine type of IIS log file to search. Accepted values are
							"W3" or "FTP".
							This parameter is only used by the script for default IIS logs
							locations
	SinceLast	(Interger)	Limit of last months logs history to analyze. Accepted values are
							Between 1 to 12.
	OutFile		(String)	Path of the desired output text file. Only full path are working.
							If this parameter is ignored, then script will generate its own
							output text file in its location directory.
							The name of the file is built as:
							
							<W3|FTP|Custom>_logs_export-<ALL|L01-L12>-<yyyyMMdd-HHmm_ss>.txt
							
							Examples:	W3_logs_export-All-20180115-1055_38.txt
										FTP_logs_export-L06-20180116-0912_19.txt

										
4. Analysis restrictions:
-------------------------

To be valid, the log file format should be in text, and data should
be separated by spaces. No other format can be accepted.

First, when the script is looking for log files, it read first line
of each log file and, in 50 first characters, looks for:
"#Software: Microsoft Internet Information Services"

If the script does not find this string, it does not consider the
log file as an IIS valid log file.

Secondly, to determine how to find the appropriate content, it read
each line, then it need to find an header line containing list of
fields or columns starting by "#Fields:" and following by columns
titles. This header line should be placed before data line to let
the script analyze logs.

Used titles are "date", "cs-uri-stem", "cs-username", "c-ip" and
"s-port". If all these fields are not in the log file the script
will no be able to analyze.
Theses fields or columns in the file can be in any order: the data
should simply be in the same order than the headers.