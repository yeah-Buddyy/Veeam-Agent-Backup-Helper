# Run as Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell.exe -Verb RunAs "-NoProfile -NoLogo -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    exit;
}

##### Edit here #####

# Default Veeam installation path
$VEEAMEndpointTray = "C:\Program Files\Veeam\Endpoint Backup\Veeam.EndPoint.Tray.exe"
$VEEAMEndpointManager = "C:\Program Files\Veeam\Endpoint Backup\Veeam.EndPoint.Manager.exe"
$VEEAMLog = "C:\ProgramData\Veeam\Endpoint\Svc.VeeamEndpointBackup.log"

# Your Veeam backup folder on the external storage with which you configured your backup job.
$USBDriveVeeamBackupPath = "E:\VeeamBackup"

# Run a new active full backup after 3 months
$NewActiveFullBackup = "3" # months

# How many full backups you want to store on your backup storage before deleting the oldest backup chain.
# If you use the default value of 3, the oldest backup chain (oldest .vbk and associated .vib files) will be deleted after 3 successful full backups.
# In this case you will always have 2 full backups and the associated incrementals (.vib) files on your backup storage.
# You will need space for at least 3 full backups, as deletion will only take place after the third successful full backup.
$DeleteBackupChain = "3" # Use minimum value of 2

# Backup no more often than 3 hours. 10800 seconds = 3 hours
$BackupNoMoreOftenThan = "10800" # Seconds

# Check every x seconds if usb drive is connected
$CheckEverySeconds = "20" # Seconds

# Path of USB_Disk_Eject.exe
$USBDiskEjectPath = "$PSScriptRoot\USB_Disk_Eject.exe"

# turn off logging the event "Could not find the drive 'E:\'. The drive might not be ready or might not be mapped" in event viewer
$LogProviderHealthEvent = $false

$systemDirectory = [System.Environment]::SystemDirectory

function Show-MessageBox {
    param (
        [string]$Message,
        [string]$Title = "Message",
        [ValidateSet("OK", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel")] [string]$Buttons = "OK",
        [ValidateSet("Critical", "Question", "Exclamation", "Information")] [string]$Icon = "Information"
    )

    # Load the necessary assembly
    Add-Type -AssemblyName Microsoft.VisualBasic

    # Map button and icon parameters to their corresponding values
    $buttonMap = @{
        OK               = 0
        OKCancel         = 1
        AbortRetryIgnore = 2
        YesNoCancel      = 3
        YesNo            = 4
        RetryCancel      = 5
    }

    $iconMap = @{
        Critical   = 16
        Question   = 32
        Exclamation = 48
        Information = 64
    }

    $buttonValue = $buttonMap[$Buttons]
    $iconValue = $iconMap[$Icon]

    # Display the message box
    [Microsoft.VisualBasic.Interaction]::MsgBox($Message, $buttonValue -bor $iconValue, $Title)
}

function Get-OldestAndSecondOldestVbkFiles {
    param (
        [string]$USBDriveVeeamJobFolderPath
    )

    # Get the oldest .vbk file path
    $oldestVbkFilePath = Get-ChildItem -Path $USBDriveVeeamJobFolderPath -Filter "*.vbk" |
        Sort-Object -Property LastWriteTime | Select-Object -First 1 -ExpandProperty FullName

    # Get the actual FileInfo object for the oldest .vbk file
    $oldestVbkFile = Get-Item -Path $oldestVbkFilePath

    # Get the second oldest .vbk file path
    $secondOldestVbkFilePath = Get-ChildItem -Path $USBDriveVeeamJobFolderPath -Filter "*.vbk" |
        Sort-Object -Property LastWriteTime | Select-Object -Skip 1 | Select-Object -First 1 -ExpandProperty FullName

    return @{
        OldestVbkFilePath = $oldestVbkFilePath
        OldestVbkFile = $oldestVbkFile
        SecondOldestVbkFilePath = $secondOldestVbkFilePath
    }
}

function Show-MessageBoxAndDeleteVibFiles {
    param (
        [string]$USBDriveVeeamJobFolderPath,
        [switch]$Force
    )

    $vbkFiles = Get-OldestAndSecondOldestVbkFiles -USBDriveVeeamJobFolderPath $USBDriveVeeamJobFolderPath
    $oldestVbkFilePath = $vbkFiles['OldestVbkFilePath']
    $oldestVbkFile = $vbkFiles['OldestVbkFile']
    $secondOldestVbkFilePath = $vbkFiles['SecondOldestVbkFilePath']

    # Get .vib files older than the second oldest .vbk file (excluding .vbm files)
    $vibFilesToDelete = Get-ChildItem -Path $USBDriveVeeamJobFolderPath -Filter "*.vib" |
        Where-Object { $_.LastWriteTime -lt (Get-Item $secondOldestVbkFilePath).LastWriteTime -and $_.Name -notlike "*.vbm" } |
        Select-Object FullName, Name, LastWriteTime

    # Add the oldest VBK file to the list for display
    $vibFilesToDelete += [PSCustomObject]@{
        FullName = $oldestVbkFilePath
        Name = (Split-Path $oldestVbkFilePath -Leaf)
        LastWriteTime = $oldestVbkFile.LastWriteTime
    }

    if ($Force) {
        # Force deletion without confirmation
        foreach ($file in $vibFilesToDelete) {
            try {
                Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                Write-Output "Deleted: $($file.FullName)"
            } catch {
                Write-Warning "Failed to delete $($file.FullName): $_"
            }
        }
    } else {
        # Display the files in a grid view
        $selectedFiles = $vibFilesToDelete | Out-GridView -Title "Select files to delete (Select All to delete the oldest backup chain)" -PassThru

        if ($selectedFiles) {
            foreach ($file in $selectedFiles) {
                try {
                    # Perform deletion
                    Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                    Write-Output "Deleted: $($file.FullName)"
                } catch {
                    Write-Warning "Failed to delete $($file.FullName): $_"
                }
            }
        } else {
            Write-Output "No files selected for deletion."
        }
    }
}

function CheckDriveHealth {
    # Define thresholds
    $MaxWearValue = 80
    $MaxRWErrors = 100
    $MaxReallocatedSectors = 1
    $MaxPendingSectors = 1

    # Get SMART failure data once
    $DriveSMARTStatuses = Get-CimInstance -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue | Where-Object { $_.PredictFailure -eq $true }

    # Obtain physical disk details
    $Disks = Get-PhysicalDisk | Where-Object { $_.BusType -match "NVMe|SATA|SAS|ATAPI|RAID" }

    # Initialize the array to accumulate all output messages
    $OutputMsgsReturn = @()

    # Loop through each disk
    foreach ($Disk in ($Disks | Sort-Object DeviceID)) {
        # Obtain disk health information
        $DiskHealth = Get-StorageReliabilityCounter -PhysicalDisk (Get-PhysicalDisk -FriendlyName $Disk.FriendlyName) | 
                      Select-Object Wear, ReadErrorsTotal, ReadErrorsUncorrected, ReadErrorsCorrected, WriteErrorsTotal, WriteErrorsUncorrected, WriteErrorscorrected, Temperature, TemperatureMax

        $DriveDetails = Get-PhysicalDisk -FriendlyName $Disk.FriendlyName | Select-Object MediaType, HealthStatus
        $DriveSMARTStatus = $DriveSMARTStatuses | Where-Object { $_.InstanceName -eq $Disk.DeviceID }

        # Calculate temperature delta
        $DiskTempDelta = $DiskHealth.Temperature - $DiskHealth.TemperatureMax

        # Create custom PSObject
        $DiskHealthState = [PSCustomObject]@{
            "Disk Number"                    = $Disk.DeviceID
            "FriendlyName"                   = $Disk.FriendlyName
            "HealthStatus"                   = $DriveDetails.HealthStatus
            "MediaType"                      = $DriveDetails.MediaType
            "Disk Wear"                      = $DiskHealth.Wear
            "Read Errors Total"              = $DiskHealth.ReadErrorsTotal
            "Temperature Delta"              = $DiskTempDelta
            "Read Errors Uncorrected"        = $DiskHealth.ReadErrorsUncorrected
            "Read Errors Corrected"          = $DiskHealth.ReadErrorsCorrected
            "Write Errors Total"             = $DiskHealth.WriteErrorsTotal
            "Write Errors Uncorrected"       = $DiskHealth.WriteErrorsUncorrected
            "Write Errors Corrected"         = $DiskHealth.WriteErrorsCorrected
            "Temperature Max"                = $DiskHealth.TemperatureMax
            "Temperature Current"            = $DiskHealth.Temperature
        }

        # Array to accumulate output messages
        $OutputMsgs = @()

        # Check conditions and set output messages
        if ($null -ne $DriveDetails.HealthStatus -and $DriveDetails.HealthStatus -ne "Healthy") {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - is in a $([string]$DriveDetails.HealthStatus.ToLower()) state"
        }

        if ($DriveSMARTStatus) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - SMART predicted failure detected with reason code $($DriveSMARTStatus.Reason)"
        }

        if ($null -ne [int]$DiskHealth.Wear -and [int]$DiskHealth.Wear -ge $MaxWearValue) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - Disk failure likely. Current wear value: $($DiskHealth.Wear), above threshold: $MaxWearValue%"
        }

        if ($null -ne [int]$DiskHealth.ReadErrorsTotal -and [int]$DiskHealth.ReadErrorsTotal -ge $MaxRWErrors) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - High number of read errors: $($DiskHealth.ReadErrorsTotal), above threshold: $MaxRWErrors"
        }

        if ($null -ne [int]$DiskHealth.WriteErrorsTotal -and [int]$DiskHealth.WriteErrorsTotal -ge $MaxRWErrors) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - High number of write errors: $($DiskHealth.WriteErrorsTotal), above threshold: $MaxRWErrors"
        }

        if ($null -ne [int]$DiskHealth.ReadErrorsCorrected -and [int]$DiskHealth.ReadErrorsCorrected -ge $MaxReallocatedSectors) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - High number of reallocated sectors: $($DiskHealth.ReadErrorsCorrected), above threshold: $MaxReallocatedSectors"
        }

        if ($null -ne [int]$DiskHealth.ReadErrorsUncorrected -and [int]$DiskHealth.ReadErrorsUncorrected -ge $MaxPendingSectors) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - High number of pending sectors: $($DiskHealth.ReadErrorsUncorrected), above threshold: $MaxPendingSectors"
        }

        if ($null -ne [int]$DiskHealth.Temperature -and [int]$DiskHealth.Temperature -gt $DiskHealth.TemperatureMax -and [int]$DiskHealth.TemperatureMax -gt 0) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - Running $DiskTempDelta degrees above max temperature: $($DiskHealth.TemperatureMax)"
        }

        if (-not $OutputMsgs) {
            $OutputMsgs += "Disk $($Disk.DeviceID) / $($Disk.FriendlyName) - is in a healthy state. No action required."
        }

        # Output the messages in color
        foreach ($msg in $OutputMsgs) {
            if ($msg -match "is in a healthy state") {
                Write-Host $msg -ForegroundColor Green
            } else {
                Write-Host $msg -ForegroundColor Red
                $OutputMsgsReturn += $msg
            }
        }
    }

    if ($OutputMsgsReturn.count -gt 0) {
        # $OutputMsgsReturn | ForEach-Object { Write-Host $_ }
        # Combine all messages into a single string
        $AllMessages = $OutputMsgsReturn -join "`n"

        $result = Show-MessageBox -Message "$AllMessages" -Title "Unhealthy drive detected, do you still want to proceed?" -Buttons "YesNoCancel" -Icon "Critical"

        # Process user response
        switch ($result) {
            'Yes' {
                break
            }
            'No' {
                Write-Verbose "User double checks their hard drives." -Verbose
                exit
            }
            'Cancel' {
                Write-Verbose "User double checks their hard drives." -Verbose
                exit
            }
        }
    }
}

# Prevent Windows from going to sleep during backup
$global:monitorScriptContent = @'
param (
    [int]$MainScriptProcessId,
    [ValidateSet('Away', 'Display', 'System')]$Option = 'System'
)

$Code=@"
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
public static extern void SetThreadExecutionState(uint esFlags);
"@

$ste = Add-Type -memberDefinition $Code -name System -namespace Win32 -passThru

# Requests that the other EXECUTION_STATE flags set remain in effect until
# SetThreadExecutionState is called again with the ES_CONTINUOUS flag set and
# one of the other EXECUTION_STATE flags cleared.
# The thread that turns it on must be the same thread that turns it off !.
$ES_CONTINUOUS = [uint32]"0x80000000"
$ES_AWAYMODE_REQUIRED = [uint32]"0x00000040"
$ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
$ES_SYSTEM_REQUIRED = [uint32]"0x00000001"

Switch ($Option) {
    "Away"    {$Setting = $ES_AWAYMODE_REQUIRED}
    "Display" {$Setting = $ES_DISPLAY_REQUIRED}
    "System"  {$Setting = $ES_SYSTEM_REQUIRED}
}

# Monitor the main script process
try {
    $mainScriptProcess = Get-Process -Id $MainScriptProcessId -ErrorAction Stop
    while ($true) {
        Write-Verbose "Staying Awake with ``${Option}`` Option" -Verbose
        $ste::SetThreadExecutionState($ES_CONTINUOUS -bor $Setting)
        Start-Sleep -Seconds 60
        [System.Console]::Clear()
        if ($mainScriptProcess.HasExited) {
            Write-Verbose "Main script process has terminated" -Verbose
            Write-Verbose "Stopping Staying Awake" -Verbose
            Start-Sleep -Seconds 3
            $ste::SetThreadExecutionState($ES_CONTINUOUS)
            break
        }
    }
} catch {
    Write-Verbose "Main script process not found or already terminated." -Verbose
    Write-Verbose "Stopping Staying Awake" -Verbose
    Start-Sleep -Seconds 3
    $ste::SetThreadExecutionState($ES_CONTINUOUS)
}

exit
'@

function PreventSleep {
    param (
        [string]$myPid
    )

    # Save the monitor script to a temporary file
    $tempFolder = [System.IO.Path]::GetTempPath()
    $monitorScriptPath = [System.IO.Path]::Combine($tempFolder, "sleep-monitor.ps1")
    Set-Content -Path $monitorScriptPath -Value $global:monitorScriptContent -Force

    # Start the monitor script
    $mainScriptProcessId = $myPid
    if (Test-Path "$monitorScriptPath" -PathType Leaf) {
        Start-Process powershell.exe -ArgumentList "-NoProfile -NoLogo -ExecutionPolicy Bypass -File `"$monitorScriptPath`" -MainScriptProcessId $mainScriptProcessId" -WindowStyle Hidden
    }
}

function Get-DriveLetter {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Validate the path format
    if ($Path -match '^[A-Za-z]:\\') {
        # Extract the drive letter
        $driveLetter = ($Path -split ':')[0]
        return $driveLetter
    } else {
        throw "Invalid path format. The path must start with a drive letter followed by a colon and a backslash (e.g., C:\)."
    }
}

try {
    $USBDriveLetter = Get-DriveLetter -Path $USBDriveVeeamBackupPath -ErrorAction Stop
    Write-Verbose "Drive letter for '$USBDriveVeeamBackupPath' is '$USBDriveLetter'" -Verbose
} catch {
    Write-Verbose "Error processing '$USBDriveVeeamBackupPath': $_" -Verbose
    Show-MessageBox -Message "Error processing '$USBDriveVeeamBackupPath': $_" -Title "Driveletter error" -Buttons "OK" -Icon "Critical"
    pause
    exit
}

function Check-VeeamLogs {
    param (
        [string]$VEEAMLog,
        [string]$OutputPath
    )

    # Extract Veeam Errors from today
    $todaysDate = Get-Date -uformat "%d.%m.%Y"

    if ([System.IO.File]::Exists($VEEAMLog)) {
        $errorLogContent = Get-Content -Path $VEEAMLog | Select-String -Pattern "\[$todaysDate\s\d+\:\d+:\d+]\s\<\d+\>\sError"
        
        if ($errorLogContent) {
            $logFolder = "$OutputPath\Veeam-Logs"
            if (-not ([System.IO.Directory]::Exists($logFolder))) {
                New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
            }

            $dotsToMinus = $todaysDate.Replace('.', '-')
            $logFilePath = "$logFolder\$dotsToMinus.log"
            $errorLogContent | Set-Content -Path $logFilePath -Force

            return $logFilePath
        } else {
            Write-Verbose "No Veeam errors found today, which is great news. :-)" -Verbose
        }
    } else {
        Write-Verbose "Could not find Svc.VeeamEndpointBackup.log" -Verbose
    }

    return $null
}


function CheckIfEnoughSpace {
    param (
        [string]$USBDriveVeeamJobFolderPath,
        [string]$USBDriveLetter
    )

    # Get the latest .vbk file
    $LatestVbkFile = Get-ChildItem "$USBDriveVeeamJobFolderPath" -Filter "*.vbk" | Sort-Object -Descending -Property LastWriteTime | Select-Object -First 1

    if (-not $LatestVbkFile) {
        Write-Verbose "No .vbk files found in the specified directory." -Verbose
        return
    }

    $LatestVbkFileSize = $LatestVbkFile.Length
    $LatestVbkFileDate = $LatestVbkFile.LastWriteTime

    $LatestVbkAndAllVibFilesAfterLatestVbk = Get-ChildItem "$USBDriveVeeamJobFolderPath" -Recurse |
        Where-Object { ($_.LastWriteTime -ge $LatestVbkFileDate) -and ($_.Extension -eq ".vib" -or $_.Extension -eq ".vbk") } |
        Measure-Object -Property Length -Sum

    $TotalBackupSize = $LatestVbkAndAllVibFilesAfterLatestVbk.Sum

    $FreeSpace = (Get-Volume -DriveLetter $USBDriveLetter).SizeRemaining

    if ($TotalBackupSize -ge $FreeSpace) {
        Write-Verbose "Backup storage does not have enough space" -Verbose
        Show-MessageBox -Message "You do not have enough disk space. Delete the old backup and run the script again" -Title "Backup storage does not have enough space" -Buttons "OK" -Icon "Critical"
        exit
    } else {
        Write-Verbose "You have enough space on your backup storage" -Verbose
    }
}

# https://helpcenter.veeam.com/docs/agentforwindows/userguide/system_requirements.html?ver=60
function Check-FileSystemType {
    param (
        [string]$USBDriveLetter
    )

    $volume = Get-Volume -DriveLetter $USBDriveLetter
    $fileSystemType = $volume.FileSystemType

    switch ($fileSystemType) {
        "NTFS" {
            Write-Verbose "We found NTFS, brilliant" -Verbose
        }
        "ReFS" {
            Write-Verbose "We found ReFS, brilliant" -Verbose
        }
        "exFAT" {
            Write-Verbose "We found exFAT, brilliant" -Verbose
        }
        default {
            Write-Verbose "We found $fileSystemType, which is not supported" -Verbose
            Show-MessageBox -Message "Please convert the FileSystemType of your backup storage device to NTFS, ReFS or exFAT and restart the script" -Title "Bad FileSystemType" -Buttons "OK" -Icon "Critical"
            exit
        }
    }
}

function Confirm-Message {
    param (
        [string]$Message
    )
    $msgBoxInput = Show-MessageBox -Message "Would you like to backup your data now?`r`n`r`nBackup method: $Message" -Title "Backup now" -Buttons "YesNoCancel" -Icon "Information"
    switch ($msgBoxInput) {
        'Yes' {
            Check-FileSystemType -USBDriveLetter $USBDriveLetter
            return $true
        }
        'No' {
            do { 
                if (([System.IO.Directory]::Exists("$USBDriveVeeamBackupPath"))) {
                    Write-Verbose "User does not want to back up any data at the moment" -Verbose
                    Write-Verbose "External hard disk still connected. Wait for user to disconnect" -Verbose
                    Start-Sleep -Seconds $CheckEverySeconds
                    [System.Console]::Clear()
                }
            } until (-not ([System.IO.Directory]::Exists("$USBDriveVeeamBackupPath")))
            return $false
        }
        'Cancel' {
            do { 
                if (([System.IO.Directory]::Exists("$USBDriveVeeamBackupPath"))) {
                    Write-Verbose "User does not want to back up any data at the moment" -Verbose
                    Write-Verbose "External hard disk still connected. Wait for user to disconnect" -Verbose
                    Start-Sleep -Seconds $CheckEverySeconds
                    [System.Console]::Clear()
                }
            } until (-not ([System.IO.Directory]::Exists("$USBDriveVeeamBackupPath")))
            return $false
        }
    }
}

# Sleep a bit, so that the script does not run immediately after a system startup
Start-Sleep -Seconds 20

While($True){
    try{
        Write-Output "Wait until the external hard disk is connected."
        Start-Sleep -Seconds $CheckEverySeconds
        [System.Console]::Clear()

        # Check if the drive letter exists
        $driveExists = Get-PSDrive -Name $USBDriveLetter -ErrorAction SilentlyContinue

        # Output the result
        if ($driveExists) {
            Write-Verbose "Drive $USBDriveLetter exists." -Verbose
        } else {
            Write-Verbose "Drive $USBDriveLetter does not exist." -Verbose
            continue
        }

        # There can only be one Job folder in this path, otherwise the script will not work.
        $USBDriveVeeamJobFolderPath = Get-ChildItem -Path "$USBDriveVeeamBackupPath" -Filter "*Job*" -Recurse -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        $VeeamJobFolderCount = $USBDriveVeeamJobFolderPath.Count
        if ($VeeamJobFolderCount -gt 1) {
            Write-Verbose "There can only be one Veeam Job Folder" -Verbose
            Show-MessageBox -Message "There can only be one Veeam Job Folder on $USBDriveVeeamBackupPath" -Title "Too many Veeam Job Folders" -Buttons "OK" -Icon "Critical"
            exit
        } else {
            $latestVbkFile = Get-ChildItem "$USBDriveVeeamJobFolderPath" -Filter "*.vbk" | Sort-Object -Descending -Property LastWriteTime | Select-Object -First 1
        }

        ##### VEEAM FIRST ACTIVE FULLBACKUP #####
        # If the Veeam Job Folder exists and there is no Activefullback in it, it will create one. Even if the Veeam Job Folder does not exist, it will create a fullbackup.
        # https://stackoverflow.com/questions/63297132/test-path-check-if-file-exist
        if ( ([System.IO.Directory]::Exists("$USBDriveVeeamJobFolderPath")) -and (-not ([System.IO.File]::Exists("$USBDriveVeeamJobFolderPath\$latestVbkFile"))) -or ($VeeamJobFolderCount -eq 0) -and ([System.IO.Directory]::Exists("$USBDriveVeeamBackupPath")) ){
            # Store the function call in a string
            $functionCall = 'Confirm-Message -message "First Active Fullbackup"'
            # Call the function using Invoke-Expression
            $callResult = Invoke-Expression $functionCall
            #Confirm-Message
            if (-not($callResult)) {
                Write-Verbose "User did not want to do a backup, so we continue" -Verbose
                continue
            }
            CheckDriveHealth
            Write-Verbose "RUNNING FIRST VEEAM ACTIVEFULLBACKUP..." -Verbose
            Start-Process -FilePath "$VEEAMEndpointTray"
            $p = Start-Process -FilePath "$VEEAMEndpointManager" -ArgumentList "/activefull" -PassThru -WindowStyle hidden
            # Get the PID of the started process
            $processPid = $p.Id
            PreventSleep -myPid $processPid
            # Loop to check if the process is still running
            while ($p.HasExited -eq $false) {
                Start-Sleep -Seconds 1
            }
            # The process has exited, continue with the rest of the script
            Write-Verbose "Veeam Process with PID $processPid has exited" -Verbose
            # https://helpcenter.veeam.com/archive/agentforwindows/50/userguide/backup_cmd.html#monitoring-backup-job-status
            $exitCode = $p.ExitCode
            if ($exitCode -eq 0) {
                Write-Verbose "FIRST VEEAM ACTIVE FULL BACKUP SUCCESSFULLY COMPLETED!" -Verbose
                $countvbkfiles = (Get-ChildItem "$USBDriveVeeamJobFolderPath" "*.vbk" | Measure-Object).Count
                if ($countvbkfiles -ge $DeleteBackupChain -and $DeleteBackupChain -ne 1 -and $DeleteBackupChain -ne 0) {
                    $message = "Do you want to delete the oldest backup chain now that you have $DeleteBackupChain full backups on your backup storage?"
                    # Show message box
                    $result = Show-MessageBox -Message "$message" -Title "Confirm Deletion" -Buttons "YesNoCancel" -Icon "Information"

                    # Process user response
                    switch ($result) {
                        'Yes' {
                            Show-MessageBoxAndDeleteVibFiles -USBDriveVeeamJobFolderPath $USBDriveVeeamJobFolderPath
                        }
                        'No' {
                            Write-Verbose "Deletion canceled by user." -Verbose
                        }
                        'Cancel' {
                            Write-Verbose "Operation canceled." -Verbose
                        }
                    }
                }
            }
            else {
                Write-Verbose "VEEAM BACKUP ERROR! ERROR CODE $exitCode" -Verbose
                Show-MessageBox -Message "Veeam Backup Error! Please Check your Veeam Logs for more details" -Title "Veeam Backup Error!" -Buttons "OK" -Icon "Critical"
                $logFilePath = Check-VeeamLogs -VEEAMLog $VEEAMLog -OutputPath $PSScriptRoot
                if ($logFilePath) {
                    Write-Verbose "Log file created at: $logFilePath" -Verbose
                    Start-Process "$systemDirectory\notepad.exe" "$logFilePath"
                } else {
                    Write-Verbose "No errors found or log file not found." -Verbose
                    Start-Process -FilePath "$VEEAMEndpointTray"
                }
                exit
            }
            Write-Verbose "The external drive will now be ejected safely, please wait." -Verbose
            Start-Sleep -Seconds 15
            # Eject the external hard drive after ActiveFullBackUp is complete.
            Start-Process -FilePath "$USBDiskEjectPath" -ArgumentList "/removeletter $USBDriveLetter" -Wait -WindowStyle hidden #-WindowStyle Minimized
            # User can check if the backup was successful in the Veeam GUI.
            Start-Process -FilePath "$VEEAMEndpointTray"
            Write-Verbose "Sleep now for $BackupNoMoreOftenThan Seconds.." -Verbose
            # Backup no more often than x hours 
            Start-Sleep -Seconds $BackupNoMoreOftenThan
            continue
        }
        ##### VEEAM ACTIVE FULLBACKUP IF OLDER THAN "X" MONTHS #####
        # If the last Veeam Activefullback is older than "x" months, a new Activefullbackup will be created.
        elseif ( ([System.IO.Directory]::Exists("$USBDriveVeeamJobFolderPath")) -and (Test-Path "$USBDriveVeeamJobFolderPath\$latestVbkFile" -OlderThan (Get-Date).AddMonths(-$NewActiveFullBackup)) ){
            # Store the function call in a string
            $functionCall = 'Confirm-Message -message "Active Fullbackup"'
            # Call the function using Invoke-Expression
            $callResult = Invoke-Expression $functionCall
            #Confirm-Message
            if (-not($callResult)) {
                Write-Verbose "User did not want to do a backup, so we continue" -Verbose
                continue
            }
            CheckDriveHealth
            CheckIfEnoughSpace -USBDriveVeeamJobFolderPath $USBDriveVeeamJobFolderPath -USBDriveLetter $USBDriveLetter
            Write-Verbose "RUNNING VEEAM ACTIVEFULLBACKUP IF OLDER THAN X MONTHS..." -Verbose
            Start-Process -FilePath "$VEEAMEndpointTray"
            $p = Start-Process -FilePath "$VEEAMEndpointManager" -ArgumentList "/activefull" -PassThru -WindowStyle hidden
            # Get the PID of the started process
            $processPid = $p.Id
            PreventSleep -myPid $processPid
            # Loop to check if the process is still running
            while ($p.HasExited -eq $false) {
                Start-Sleep -Seconds 1
            }
            # The process has exited, continue with the rest of the script
            Write-Verbose "Veeam Process with PID $processPid has exited" -Verbose
            # https://helpcenter.veeam.com/archive/agentforwindows/50/userguide/backup_cmd.html#monitoring-backup-job-status
            $exitCode = $p.ExitCode
            if ($exitCode -eq 0) {
                Write-Verbose "SUCCESSFULLY COMPLETED VEEAM ACTIVEFULLBACKUP (IF OLDER THAN X MONTHS)" -Verbose
                $countvbkfiles = (Get-ChildItem "$USBDriveVeeamJobFolderPath" "*.vbk" | Measure-Object).Count
                if ($countvbkfiles -ge $DeleteBackupChain -and $DeleteBackupChain -ne 1 -and $DeleteBackupChain -ne 0) {
                    $message = "Do you want to delete the oldest backup chain now that you have $DeleteBackupChain full backups on your backup storage?"
                    # Show message box
                    $result = Show-MessageBox -Message "$message" -Title "Confirm Deletion" -Buttons "YesNoCancel" -Icon "Information"

                    # Process user response
                    switch ($result) {
                        'Yes' {
                            Show-MessageBoxAndDeleteVibFiles -USBDriveVeeamJobFolderPath $USBDriveVeeamJobFolderPath
                        }
                        'No' {
                            Write-Verbose "Deletion canceled by user." -Verbose
                        }
                        'Cancel' {
                            Write-Verbose "Operation canceled." -Verbose
                        }
                    }
                }
            }
            else {
                Write-Verbose "VEEAM BACKUP ERROR! ERROR CODE $exitCode" -Verbose
                Show-MessageBox -Message "Veeam Backup Error! Please Check your Veeam Logs for more details" -Title "Veeam Backup Error!" -Buttons "OK" -Icon "Critical"
                $logFilePath = Check-VeeamLogs -VEEAMLog $VEEAMLog -OutputPath $PSScriptRoot
                if ($logFilePath) {
                    Write-Verbose "Log file created at: $logFilePath" -Verbose
                    Start-Process "$systemDirectory\notepad.exe" "$logFilePath"
                } else {
                    Write-Verbose "No errors found or log file not found." -Verbose
                    Start-Process -FilePath "$VEEAMEndpointTray"
                }
                exit
            }
            Write-Verbose "The external drive will now be ejected safely, please wait." -Verbose
            Start-Sleep -Seconds 15
            # Eject the external hard drive after ActiveFullBackUp is complete.
            Start-Process -FilePath "$USBDiskEjectPath" -ArgumentList "/removeletter $USBDriveLetter" -Wait -WindowStyle hidden #-WindowStyle Minimized
            # User can check if the backup was successful in the Veeam GUI.
            Start-Process -FilePath "$VEEAMEndpointTray"
            Write-Verbose "Sleep now for $BackupNoMoreOftenThan Seconds.." -Verbose
            # Backup no more often than x hours 
            Start-Sleep -Seconds $BackupNoMoreOftenThan
            continue
        }
        else{
            ##### VEEAM INCREMENTAL BACKUP #####
            # If the last Veeam Activefullback is not older than "X" months, it will only create an incremental backup.
            if ( ([System.IO.Directory]::Exists("$USBDriveVeeamJobFolderPath")) -and ([System.IO.File]::Exists("$USBDriveVeeamJobFolderPath\$latestVbkFile")) ){
                # Store the function call in a string
                $functionCall = 'Confirm-Message -message "Incremental Backup"'
                # Call the function using Invoke-Expression
                $callResult = Invoke-Expression $functionCall
                #Confirm-Message
                if (-not($callResult)) {
                    Write-Verbose "User did not want to do a backup, so we continue" -Verbose
                    continue
                }
                CheckDriveHealth
                Write-Verbose "RUNNING VEEAM INCREMENTAL BACKUP..." -Verbose
                Start-Process -FilePath "$VEEAMEndpointTray"
                $p = Start-Process -FilePath "$VEEAMEndpointManager" -ArgumentList "/backup" -PassThru -WindowStyle hidden
                # Get the PID of the started process
                $processPid = $p.Id
                PreventSleep -myPid $processPid
                # Loop to check if the process is still running
                while ($p.HasExited -eq $false) {
                    Start-Sleep -Seconds 1
                }
                # The process has exited, continue with the rest of the script
                Write-Verbose "Veeam Process with PID $processPid has exited" -Verbose
                # https://helpcenter.veeam.com/archive/agentforwindows/50/userguide/backup_cmd.html#monitoring-backup-job-status
                $exitCode = $p.ExitCode
                if ($exitCode -eq 0) {
                    Write-Verbose "SUCCESSFULLY COMPLETED VEEAM INCREMENTAL BACKUP!" -Verbose
                }
                else {
                    Write-Verbose "VEEAM BACKUP ERROR! ERROR CODE $exitCode" -Verbose
                    Show-MessageBox -Message "Veeam Backup Error! Please Check your Veeam Logs for more details" -Title "Veeam Backup Error!" -Buttons "OK" -Icon "Critical"
                    $logFilePath = Check-VeeamLogs -VEEAMLog $VEEAMLog -OutputPath $PSScriptRoot
                    if ($logFilePath) {
                        Write-Verbose "Log file created at: $logFilePath" -Verbose
                        Start-Process "$systemDirectory\notepad.exe" "$logFilePath"
                    } else {
                        Write-Verbose "No errors found or log file not found." -Verbose
                        Start-Process -FilePath "$VEEAMEndpointTray"
                    }
                    exit
                }
                Write-Verbose "The external drive will now be ejected safely, please wait." -Verbose
                Start-Sleep -Seconds 15
                # Eject the external hard drive after ActiveFullBackUp is complete.
                Start-Process -FilePath "$USBDiskEjectPath" -ArgumentList "/removeletter $USBDriveLetter" -Wait -WindowStyle hidden #-WindowStyle Minimized
                # User can check if the backup was successful in the Veeam GUI.
                Start-Process -FilePath "$VEEAMEndpointTray"
                Write-Verbose "Sleep now for $BackupNoMoreOftenThan Seconds.." -Verbose
                # Backup no more often than x hours 
                Start-Sleep -Seconds $BackupNoMoreOftenThan
                continue
            }
        }
        
    }
    catch [System.Management.Automation.DriveNotFoundException]{
        Write-Host -ForegroundColor Red "Drive not found, continue searching.."
        # Write-Host -ForegroundColor Yellow $Error[0].Exception.GetType()
    }
    catch{
       Write-Host -ForegroundColor Yellow "General Exception"
       Write-Host -ForegroundColor Magenta $Error[0].Exception
    }
}
