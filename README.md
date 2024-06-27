# Readme
If you leave the configuration in the Powershell script as it is, a full backup is created every 3 months. Everything in between is backed up incrementally. After 3 successful full backups, the oldest full backup and the associated incremental backups are deleted. So you always have 2 full backups + the incremental ones and can go back 6 months. Personally, this is enough for me, but you can customise this in the script.

If you want to use the script, you have to adjust the path to your Veeam installation location and the path of the backup folder that you specified when creating the Veeam job in the script.
You can then either start the script manually as soon as you have connected your external hard drive or you can create a task in the task scheduler with the "Create-Task.bat" script. The backup will then always start automatically as soon as you have connected the hard drive.

When creating the Veeam job, please also note the following:


# Ensure that "Job" is included in the name

![jobname](https://github.com/yeah-Buddyy/Veeam-Agent-Backup-Helper/assets/170236793/e4d19d13-2efc-4294-9e3e-1e5590ef7961)

# Set "Keep backups for" to maximum (730)

![keepbackupsfor](https://github.com/yeah-Buddyy/Veeam-Agent-Backup-Helper/assets/170236793/2f202855-963e-4ba4-91ff-45dfe28912ae)

# Do not tick "Create active full backups periodically"

![createactivefull](https://github.com/yeah-Buddyy/Veeam-Agent-Backup-Helper/assets/170236793/546ae0b7-c813-4e07-b7eb-888ce3feaeab)


# Do not tick anything under Schedule

![schedule](https://github.com/yeah-Buddyy/Veeam-Agent-Backup-Helper/assets/170236793/cbbdbc0d-1b23-4fc0-ae6a-3c2458860b8d)

# USB_Disk_Eject.exe Original Download
https://github.com/bgbennyboy/USB-Disk-Ejector 
