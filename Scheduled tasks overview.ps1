<#
.SYNOPSIS
    Send a mail with all scheduled tasks in attachment.

.DESCRIPTION
    Collect a list of all scheduled tasks with state 'Enabled'. Send this 
    list by e-mail to the users. This can be useful as an overview for the 
    management. 

.PARAMETER TaskPath
    The folder in the Task Scheduler in which the tasks are stored.

.PARAMETER MailTo
    List of e-mail addresses where the e-mail will be sent.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$TaskPath,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\Windows task scheduler\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}
Process {
    Try {
        #region Get scheduled tasks
        $M = "Get scheduled tasks in folder '$TaskPath'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $tasks = Get-ScheduledTask -TaskPath "\$TaskPath\*"
        #endregion

        #region Filter only Enabled tasks
        $tasksToExport = $tasks | Where-Object State -NE Disabled
        
        $M = "Enabled scheduled tasks: $($tasksToExport.Count)"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$($tasksToExport.Count) scheduled tasks in '$TaskPath'"
            Message   = "<p>A total of <b>$($tasksToExport.Count) scheduled tasks</b> with state 'Enabled' have been found in folder '$($TaskPath)'.</p>"
            LogFolder = $logParams.LogFolder
            Header    = $ScriptName
            Save      = "$logFile - Mail.html"
        }

        #region Export tasks to Excel file
        if ($tasksToExport) {
            $tasksToExport | ForEach-Object {
                $M = "TaskName '{0}' TaskPath '{1}' State '{2}'" -f
                $($_.TaskName), $($_.TaskPath), $($_.State)
                Write-Verbose $M
            }
            
            $excelParams = @{
                Path          = "$LogFile - Overview.xlsx"
                AutoSize      = $true
                FreezeTopRow  = $true
                WorkSheetName = 'Tasks'
                TableName     = 'Tasks'
            }

            $M = "Export {0} scheduled tasks to Excel file '{1}'" -f
            $($tasksToExport.Count), $excelParams.Path
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $tasksToExport |
            Select-Object -Property TaskName, TaskPath, State, Description |
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
            $mailParams.Message += "<p><i>* Check the attachment for details</i></p>"
        }
        #endregion

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}