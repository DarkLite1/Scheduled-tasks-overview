#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Function New-TaskObjectsHC {
        Param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [HashTable[]]$Hash
        )
    
        Process {
            $params = @{
                TypeName     = 'Microsoft.Management.Infrastructure.CimInstance' 
                ArgumentList = @('MSFT_ScheduledTask')
            }
            foreach ($H in $Hash) {
                $Obj = New-Object @params                 
                $H.GetEnumerator() | ForEach-Object {
                    $Obj.CimInstanceProperties.Add([Microsoft.Management.Infrastructure.CimProperty]::Create($_.Key, $_.Value, [Microsoft.Management.Infrastructure.CimFlags]::None))  
                }
                $Obj
            }
        }
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test'
        TaskPath   = 'contosoTasks'
        MailTo     = 'bob@contoso.com'
        LogFolder  = (New-Item -Path 'TestDrive:\Log' -ItemType Directory).FullName
    }

    Mock Get-ScheduledTask
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'the mandatory parameters are' {
    It "<_>" -ForEach @('ScriptName', 'TaskPath', 'MailTo') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | Should -BeTrue
    }
}

Describe 'when the script runs' {
    BeforeAll {
        $testTasks = @(
            @{
                TaskPath    = "\contosoTasks\task1"
                TaskName    = 'task 1'
                State       = 'Running'
                Description = 'test'
            }
            @{
                TaskPath    = "\contosoTasks\task2"
                TaskName    = 'task 2'
                State       = 'Ready'
                Description = 'test'
            }
            @{
                TaskPath    = "\contosoTasks\task3"
                TaskName    = 'task 3'
                State       = 'Disabled'
                Description = 'test'
            }
        )

        Mock Get-ScheduledTask {
            $testTasks | New-TaskObjectsHC
        }

        .$testScript @testParams
    }
    It "The log folder 'Scheduled tasks' is created" {
        'TestDrive:\Log\Scheduled tasks' | Should -Exist
    }
    It 'Get-ScheduledTask is called' {
        Should -Invoke Get-ScheduledTask -Times 1 -Exactly -Scope Describe -ParameterFilter {
            $TaskPath -eq "\contosoTasks\*"
        }
    }
    It 'All tasks are collected' {
        $tasks.Count | Should -BeExactly 3
    }
    It 'Disabled tasks are ignored for export' {
        $tasksToExport.Count | Should -BeExactly 2
        $tasksToExport.State | Should -Not -Contain Disabled
    }
    Context 'An Excel file is created' {
        BeforeAll {
            $testExportedExcelFile = (Get-ChildItem 'TestDrive:\Log\Scheduled tasks\Test' -Filter '*.xlsx').FullName
        }
        It 'in the log folder' {
            $testExportedExcelFile | Should -Exist
        }
        It 'containing the enabled scheduled tasks' {
            $testData = Import-Excel -Path $testExportedExcelFile
            $testData | Should  -HaveCount 2
            $testData[0].TaskName | Should  -Be 'task 1'
            $testData[1].TaskName | Should  -Be 'task 2'
        }
    }
    Context 'an e-mail is sent to the user' {
        It 'with the Excel file in attachment' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Subject -eq "2 scheduled tasks in 'contosoTasks'") -and
                ($Message -like "*2 scheduled tasks*Enabled*") -and
                ($LogFolder) -and
                ($Header -eq $ScriptName) -and
                ($Save -like '* - Mail.html') -and
                ($EmailParams.Attachments -like '*.xlsx')
            }
        }
    }
}