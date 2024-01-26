<#
.SYNOPSIS
    This script performs backup of virtual machines on XenServer or Citrix Hypervisor.

.DESCRIPTION
    The script connects to the specified XenServer or Citrix Hypervisor host and performs backup of the specified virtual machines.
    It creates a snapshot of each virtual machine, exports the snapshot to a specified backup location, and removes the snapshot.
    After the backup process is completed, it sends an email with the log file as an attachment.

.PARAMETER XenServerHost
    The URL of the XenServer or Citrix Hypervisor host.

.PARAMETER Username
    The username for authentication.

.PARAMETER Password
    The password for authentication.

.PARAMETER VMNames
    The name of the virtual machine(s) to backup.

.PARAMETER BackupDir
    The backup location.

.PARAMETER LogDir
    The directory to store log files.

.PARAMETER AdminEmail
    The email address of the administrator.

.PARAMETER ToEmail
    The email address of the recipient.

.PARAMETER SmtpServer
    The SMTP server for sending emails.

.PARAMETER SmtpPort
    The SMTP port for sending emails.

.PARAMETER HypervisorVersion
    The version of the hypervisor (1 for XenServer, 2 for Citrix Hypervisor).

.EXAMPLE
    .\BackupScript.ps1 -XenServerHost "xenserver.example.com" -Username "root" -Password "P@ssw0rd123" -VMNames "VM01" -BackupDir "C:\Temp" -LogDir "C:\Logs" -AdminEmail "admin@example.com" -ToEmail "recipient@example.com" -SmtpServer "smtp.office365.com" -SmtpPort "587" -HypervisorVersion "2"

.NOTES
    This script must run with Powershell 5 for Citrix Hypervisor and Powershell 7 for XenServer.
    For more information, see https://www.newyard.nl.
    Version: 0.1
    Author: Sjoerd van den Nieuwenhof
    Company: New Yard
#>

# Set the hypervisor version (1 for XenServer, 2 for Citrix Hypervisor)
$HypervisorVersion = "1"

# Set the XenServer URL
$XenServerHost = "xenserver.example.com"

# Set the username for authentication
$Username = "root"

# Set the password for authentication
$Password = "P@ssw0rd123"

# Set the name of the virtual machine
$VMNames = "VM01"   

# Get the current date in the format "yyyy-MM-dd"
$date = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"

# Set the name for the backup
$BackupName = "Backup"

# Set the name for the snapshot using the backup name and date
$SnapshotName = "$BackupName-$date"

# Set the backup location
$BackupDir  = "C:\Temp"
$logDir = $BackupDir

# Get the current location
$source = Get-Location

# Start logging
$datelogfile = Get-Date -format "yyyy-dd-MM"
$logfile = "$LogDir\XenServerBackup_$datelogfile.txt"

Write-Host "$Date --- Start Backup Process"
Write-Output "$Date --- Start Backup Process for $VMNames" `n | Out-File -Append $logfile

#Import the XenServer module
Import-Module XenServerPSModule
Write-Host "XenServer Module imported"

#Connect to XenServer
Write-Host "Connecting to XenServer"
try {
    Connect-XenServer -url https://$XenServerHost -username $Username -password $Password -NoWarnCertificates -SetDefaultSession
    Write-Host "Connected to XenServer"
}
catch {
    #When the connection fails, get the PoolMaster and connect to that
    #[1] is the PoolMaster ip displayed in the error message
    Write-Host "Failed to connect to XenServer. Error: $($_.Exception.ErrorDescription[1])"
    $PoolMaster = $_.Exception.ErrorDescription[1]
    #connect to the PoolMaster
    Connect-XenServer -url https://$PoolMaster -username $Username -password $Password -NoWarnCertificates -SetDefaultSession
    Write-Host "Connected to XenServer Poolmaster"
}

    # Iterate over each VM name
    foreach ($VMName in $VMNames) {
        $dateStarted = Get-Date -Format "HH:mm:ss"
        write-Output "$dateStarted --- Start Backup Process for $VMName to $BackupDir\$vmname" | Out-File -Append $logfile
        Write-Output "Getting VM: $VMName" | Out-File -Append $logfile
        if (Get-XenVM -Name $VMName) {
            Write-Output "$vmname VM exists" | Out-File -Append $logfile
            # Check if the folder exists in the backup location per VM
            $folderPath = Join-Path -Path $BackupDir -ChildPath $VMName
            if (!(Test-Path -Path $folderPath)) {
                # Create the folder
                New-Item -ItemType Directory -Path $folderPath | Out-Null
                Write-Output "Folder $VMName created in $BackupDir" | Out-File -Append $logfile
            } else {
                Write-Output "Folder $VMName already exists in $BackupDir" | Out-File -Append $logfile
            }
            # Set the backup location per VM
            $BackupDirPerVM = Join-Path -Path $BackupDir -ChildPath $VMName

            $SnapshotNamePerVm = "$VMName-$SnapshotName"
            
            Write-Output "Creating Snapshot $SnapshotNamePerVm" | Out-File -Append $logfile
            # Create a snapshot with the name $SnapshotName
            Invoke-XenVM -Name $VMName -XenAction Snapshot -NewName $SnapshotNamePerVm | Out-File -Append $logfile

            Write-Output "Exporting $SnapshotNamePerVm Snapshot" | Out-File -Append $logfile
            # Get the uuid of the snapshot
            $backupuuid = Get-XenVM -Name $SnapshotNamePerVm
            if ($backupuuid) {
                try {
                    # Export the snapshot to $BackupDir\$SnapshotName.xva in XVA format
                    Write-Output "Exporting Snapshot to $BackupDirPerVM\$SnapshotNamePerVm" | Out-File -Append $logfile
                    Export-XenVm -Uuid $backupuuid.uuid -XenHost $XenServerHost -Path "$BackupDirPerVM\$SnapshotNamePerVm.xva"
                } catch {
                    Write-Output "Failed to export snapshot: $($_.Exception.Message)" | Out-File -Append $logfile
                }
                try {
                    Write-Output "Removing Snapshot $SnapshotNamePerVm" | Out-File -Append $logfile
                    Remove-XenVM -Uuid $backupuuid.uuid
                } catch {
                    Write-Output "Failed to remove snapshot: $($_.Exception.Message)" | Out-File -Append $logfile
                }
            } else {
                Write-Output "Snapshot not found" | Out-File -Append $logfile
            }
        } else {
            Write-Output "VM does not exist" | Out-File -Append $logfile
        }
        $dateFinished = Get-Date -Format "HH:mm:ss"
        Write-Output "$dateFinished --- Finished Backup Process for $VMName to $BackupDir" `n | Out-File -Append $logfile
        
    }

#Disconnect from XenServer to close the session
Disconnect-XenServer


if ($HypervisorVersion -eq "1") {
    # Set the subject and body for the email to XenServer
    $subject = "XenServer Backup"
    $body = "XenServer Backup completed"
} else {
    # Set the subject and body for the email to Citrix hypervisor
    $subject = "Citrix Hypervisor Backup"
    $body = "Citrix Hypervisor Backup completed"
}

$secretValueText = "MySecretPassword123"
$AdminEmail = "admin@example.com"
$ToEmail = "recipient@example.com"
$SmtpServer = "smtp.office365.com"
$SmtpPort = "587"
$subjectDate = "$subject $datelogfile"

#creating Credentials
Write-Output "Creating Credentials"
$SecurePassword = ConvertTo-SecureString $secretValueText -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminEmail, $SecurePassword

function Send-LogfileMail {
    param (
        [string]$Subject,
        [string]$Body,
        [string]$To,
        [string]$Attachments
    )

    Send-MailMessage -From $AdminEmail -To $ToEmail -Subject $Subject -Body $Body -Attachments $Attachments -SmtpServer $SmtpServer -Credential $Credential -UseSsl -Port $SmtpPort
}

Send-LogfileMail -Subject $SubjectDate -Body $Body -To $ToEmail -Attachments $logfile
