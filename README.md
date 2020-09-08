# ConfigMgr-LogFile-Opener
![Alt Text](https://msitproblog.com/wp-content/uploads/2020/04/configmgr_logfile_opener_2.1.2_part1.png)

## Description
This Tool automates the usage of CMTrace, CMLogViewer and OneTrace for opening single or multiple ConfigMgr Client LogFiles. Besides handling LogFiles, the Tool can be used to execute the most common ConfigMgr Client actions.

The Full Description and Usage Documentation can be found on my blog: http://msitproblog.com/2016/11/20/configmgr-logfile-opener-released/

## Requirements
Starting with Version 1.2.0, ConfigMgr LogFile Opener requires PowerShell 3.0 or higher.

## Parameters
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Parameter&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description
------------ | -------------
`-Hostname` |  Can be used for a direct connection to a client device. Otherwise the Tool will prompt you to specify the Hostname.
`-CMTrace` |  Can be used to specify a different location for CMLogViewer.exe. The Tool will look by default at “C:\Program Files (x86)\Configuration Manager Support Center\CMLogViewer.exe”
`-CMLogViewer` |  Can be used to specify a different location for CMTrace.exe. The Tool will look by default at "C:\Windows\CMTrace.exe".
`-ClientLogFilesDir` |  Can be used to specify a different location for the ConfigMgr Client LogFiles. e.g. 'c$\Program Files\CCM\Logs'.
`-DisableLogFileMerging` |  Can be used to prevent CMTrace from merging multiple LogFiles. The 'Window' Menu has to be used to toggle between the different LogFiles.
`-WindowStyle` |  Can be used to change the Window Mode of CMTrace and File Explorer. Possible values are 'Minimized', 'Maximized' and 'Normal'.
`-CMTraceActionDelay` |  Specify the amount of time in milliseconds, the Script should wait between the Steps when opening multiple LogFiles in CMTrace. Default value is 1500
`-ActiveLogProgram` |  Specify which Log Program (CMTrace or CMLogViewer) should be active when the tool is starting. Default value is ‘CMTrace’
`-DisableHistoryLogFiles` |  If specified, the Tool won’t open any history log files
`-DisableUpdater` |  If specified, the Tool won't prompt if there is a newer Version available
`-EnableAutoLogLaunch` |  If specified, the Tool will automatically open the corresponding logs when executing client actions.
