**SFTP_Via_WinSCP**

This script can be configured to monitor a folder and SFTP any files to a SFTP server. It will not same the same file twice. This script is preferably ran through Windows Task Scedhuler; however, can could also use a `while ($true)` loop to continuously have the script run. 

This script requires WinSCP to be installed wherever the script is ran.  

Requirements: 
1. Your system must have "C:\Program Files (x86)\WinSCP\WinSCPnet.dll". If this .dll is located elsewhere, simply update the PowerShell script to reflect the appropriate path:
	Import-Module "%yourPathHere%"

2. You must define the following script variables to fit your needs: 
	$WinSCP_SessionOptions 
	$sftp_paths 

3. This script requires that you encrypt your SFTP credentials to an credential file running the command below: 
	Get-Credential | Export-Clixml -path "C:\creds.txt" -Verbose 
