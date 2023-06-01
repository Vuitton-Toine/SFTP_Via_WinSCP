# This script requires WinSCP to be installed wherever the script is ran. It must have C:\Program Files (x86)\WinSCP\WinSCPnet.dll
#      Script created by Anthony Mignona, 05/05/2023

Import-Module "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

function Send-FilesViaSFTP {
    param (
        $WinSCPSessionOptions,
        $PathsObject,
        $FilesToSend
    )

    # Secure Credentials. Works in conjunction with the steps outlined in commented lines 1-4. 
    $login = Import-Clixml -path $PathsObject.credential_file
    $plogin = $login.password 
    $WinSCPSessionOptions.Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($plogin))
    $WinSCPSessionOptions.Username = $login.UserName
   
    if($FilesToSend){
        write-host "opening connection"
        $session = New-Object WinSCP.Session
        $session.Open($WinSCPSessionOptions)
    }

    foreach($FileToSend in $FilesToSend){
        write-host $FileToSend -BackgroundColor DarkGreen -ForegroundColor White
        
        # if "filename + transfer successful" is in the file, skip it.
        $pattern = $FileToSend.Name.tostring()
        $pattern = [regex]::Escape($pattern)
        if(Select-string -Path $PathsObject.log_file -Pattern ("File " + $pattern + " transfer successful") ){
            write-host "Detected in log file" 
            continue
        }

        # Set local and remote paths
        $localPath = $FileToSend.FullName
        $remotePath = ($PathsObject.external_destination_path + $FileToSend.Name)
                
        # Transfer Options (Only required if you have a requirement not to send partial files)
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::off   # required to not send partial files. 
 
        # Upload the file to the remote server.
        $transferResult = $session.PutFiles($localPath, $remotePath, $False, $transferOptions) 

        # Check if the transfer was successful
        if ($transferResult.IsSuccess -eq $true) {
            # If the transfer was successful, move the file to another location locally
            Write-Host "File $($FileToSend.Name) transfer successful" -BackgroundColor Green -ForegroundColor White
            $log_content = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - File $($FileToSend.Name) transfer successful" 
            Move-Item $localPath $PathsObject.sent_archive_path -force -Verbose

        } elseif($transferResult.IsSuccess -eq $false){
            Write-Host "File $($FileToSend.Name) transfer failed" -BackgroundColor Red -ForegroundColor White
            $log_content = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - File $($FileToSend.Name) transfer failed" 
        }else{
            write-host "do nothing" 
        }

        Add-Content -Path $PathsObject.log_file.ToString() -Value $log_content -Verbose
        $log_content = get-content -Path $PathsObject.log_file | Select-Object -Last 5000
        Set-Content -Path $PathsObject.log_file.ToString() -Value $log_content -Verbose
    }

    if($session.Opened){
        write-host "closing connection"
        # Close the session
        $session.Dispose()
    }
}

function CheckIf-FileSent{
    param (
        $PathsObject
    )

    foreach($File in (Get-ChildItem -Path $PathsObject.internal_source_path -File)) {
        $pattern = $File.Name.tostring()
        $pattern = [regex]::Escape($pattern)
        if(Select-string -Path $PathsObject.log_file -Pattern ("File "+ $pattern + " transfer successful")){     
            continue
        }
        Copy-Item -Path $File.FullName -Destination $PathsObject.internal_destination_path -Force -Verbose
    }
}

$WinSCP_SessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = "" # SFTP Server FQDN should be here. For example: "test.company.com"
    UserName = "" # This can stay blank. Be sure to update the $sftp_paths.credential_file attribute
    Password = "" # This can stay blank. Be sure to update the $sftp_paths.credential_file attribute
    PortNumber =  # What is the port number that you need to connect to? 
    SshHostKeyFingerprint = "" # This can be obtained from WinSCP GUI once you connect to the SFTP server in question. For example: "ssh-rsa 2048 xxxxxxxxxxxxxxxxxxxxx/xxxxxxxxxxxxxxxxxxxxx"
}

$sftp_paths = [PSCustomObject]@{
    internal_source_path     = "" # if files originate elsewhere within the network, put that folder path here. 
    internal_destination_path = "" # Where should files queue up for sending? 
    sent_archive_path = "" # Where do you want your files to be stored locally after they're sent? 
    external_destination_path = "" # Which folder are you sending to once you connect to the SFTP? For instance, /test/
    log_file = "" # Where do you want your log file so that you can see successes/failures? 
    credential_file = "" # Where did you export your encrypted credentials to the SFTP server? For instance: get-credential | Export-Clixml -path "C:\creds.txt" -Verbose  
}

# Check if the files were already sent, if they werent, queue them up for sending.. 
CheckIf-FileSent -PathsObject $sftp_paths

# Grab the files in Queue
$files_to_send = Get-ChildItem -Path $sftp_paths.internal_destination_path -filter "*.*"

# Send the files in Queue
Send-FilesViaSFTP -WinSCPSessionOptions $WinSCP_SessionOptions -PathsObject $sftp_paths -FilesToSend $files_to_send

# Delete sent files > 180 days
Get-ChildItem -Path $sftp_paths.sent_archive_path -Filter *.* | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-180) } | Remove-Item -Verbose
