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
    [String[]]$MailTo = @(),
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\Windows task scheduler\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
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
        $M = "Get scheduled tasks in folder '$TaskPath'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $tasks = Get-ScheduledTask -TaskPath "\$TaskPath\*"
    
        $tasksToExport = $tasks | Where-Object State -NE Disabled
        
        $M = "Enabled scheduled tasks: $($tasksToExport.Count)"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $emailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$($tasksToExport.Count) scheduled tasks in '$TaskPath'"
            Message   = "<p>A total of <b>$($tasksToExport.Count) scheduled tasks</b> with state 'Enabled' have been found in folder '$($TaskPath)'.</p>"
            LogFolder = $logParams.LogFolder
            Header    = $ScriptName
            Save      = $logFile + ' - Mail.html'
        }

        if ($tasksToExport) {
            Foreach ($task in $tasksToExport) {
                $M = "TaskName '{0}' TaskPath '{1}' State '{2}'" -f
                $($task.TaskName), $($task.TaskPath), $($task.State)
                Write-Verbose $M
            }
            
            $M = "Export $($tasksToExport.Count) scheduled tasks to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $excelParams = @{
                Path          = $LogFile + '.xlsx'
                AutoSize      = $true
                FreezeTopRow  = $true
                WorkSheetName = 'Tasks'
                TableName     = 'Tasks'
            }
            $tasksToExport | Select-Object TaskName, TaskPath, State, Description | 
            Export-Excel @excelParams

            $emailParams.Attachments = $excelParams.Path
            $emailParams.Message += "<p><i>* Check the attachment for details</i></p>"
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @emailParams
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