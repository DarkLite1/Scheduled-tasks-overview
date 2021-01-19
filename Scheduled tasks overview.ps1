<#
.SYNOPSIS
    Send a mail with all scheduled tasks in attachment.

.DESCRIPTION
    Collect a list of all scheduled tasks with state 'Enabled'. Send this 
    list by e-mail to the users. This can be useful as an overview for the management. 

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
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'
)

Begin {
    try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        $logParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath "Scheduled tasks\\$ScriptName"
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $logFile = New-LogFileNameHC @LogParams
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
        $tasks = Get-ScheduledTask -TaskPath "\$TaskPath\*"
        Write-Verbose "Retrieved $($tasks.Count) tasks in folder '$TaskPath'"
    
        $tasksToExport = $tasks | Where-Object State -NE Disabled
        Write-EventLog @EventVerboseParams -Message "Enabled scheduled tasks: $($tasksToExport.Count)"

        $emailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$($tasksToExport.Count) scheduled tasks"
            Message   = "<p>A total of <b>$($tasksToExport.Count) scheduled tasks</b> with state 'Enabled' have been found in folder '$($TaskPath)'.</p>"
            LogFolder = $logParams.LogFolder
            Header    = $ScriptName
            Save      = $logFile + ' - Mail.html'
        }

        if ($tasksToExport) {
            Foreach ($task in $tasksToExport) {
                Write-Verbose "TaskName '$($task.TaskName)' TaskPath '$($task.TaskPath)' State '$($task.State)'"
            }
            Write-EventLog @EventOutParams -Message "Export $($tasksToExport.Count) scheduled tasks to Excel"

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