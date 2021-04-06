# ConfigMgr-LogFile-Opener
![Alt Text](https://msitproblog.com/wp-content/uploads/2021/04/configmgr_logfile_opener_3.0.1_part1.jpg)
![Alt Text](https://msitproblog.com/wp-content/uploads/2021/04/configmgr_logfile_opener_3.0.1_part2.jpg)
![Alt Text](https://msitproblog.com/wp-content/uploads/2021/04/configmgr_logfile_opener_3.0.1_part3.jpg)

## Description
This Tool automates the usage of CMTrace, CMLogViewer and OneTrace for opening single or multiple ConfigMgr Client LogFiles. Besides handling LogFiles, the Tool can be used to execute the most common ConfigMgr Client actions.

The Full Description and Usage Documentation can be found on my blog: https://msitproblog.com/configmgr-logfile-opener

## Requirements
* Starting with Version 3.0.0, Windows 8.1 / Server 2012 R2 or later is required
* Starting with Version 1.2.0, PowerShell 3.0 or later is required.

## Parameters
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Parameter&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description
------------ | -------------
`-Hostname` |  Can be used for a direct connection to a client device. Otherwise the Tool will prompt you to specify the Hostname
`-CMTrace` |  Can be used to specify a different location for CMTrace.exe. The Tool will look by default at "C:\Windows\CMTrace.exe"
`-CMLogViewer` |  Can be used to specify a different location for CMLogViewer.exe. The Tool will look by default at “C:\Program Files (x86)\Configuration Manager Support Center\CMLogViewer.exe”
`-OneTrace` |  Can be used to specify a different location for CMPowerLogViewer.exe. The Tool will look by default at "C:\Program Files (x86)\Configuration Manager Support Center\CMPowerLogViewer.exe"
`-ClientLogFilesDir` |  Can be used to specify a different location for the ConfigMgr Client LogFiles. e.g. 'c$\Program Files\CCM\Logs'
`-ActionDelayShort` |  Can be used to specify the amount of time in milliseconds, the Script should wait between the Steps when opening multiple LogFiles in GUI Mode. Default value is 1500
`-ActionDelayLong` |  Can be used to specify the amount of time in milliseconds, the Script should wait between the Steps when opening multiple LogFiles in GUI Mode. Default value is 2500
`-LogProgram` | Can be used to specify which Log Program should be active when the tool is starting. Possible values are "CMTrace, "CMLogViewer" and "OneTrace". Default value is 'CMTrace'
`-LogProgramWindowStyle` |  Can be used to specify the WindowStyle of CMTrace and File Explorer. Possible values are "Minimied", "Maximized" and "Normal". Default value is 'Normal'"
`-DisableHistoryLogFiles` |  If specified, the Tool won’t open any history log files
`-RecentLogLimit` |  Can be used to specify the number of recent log files which will be listed in the menu. Default value is 15"
`-DisableUpdater` |  If specified, the Tool won't prompt if there is a newer Version available
`-EnableAutoLogLaunch` |  If specified, the Tool will automatically open the corresponding logs when executing client actions
