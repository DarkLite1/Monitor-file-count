#Requires -Version 7
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Send an e-mail when there are more than x files in a folder.

    .DESCRIPTION
        Count all the files in a folder and send an e-mail when there are more
        than x files in the specified folder.

    .PARAMETER ImportFile
        Contains all the parameters for the script.

    .PARAMETER PSSessionConfiguration
        The version of PowerShell on the remote endpoint as returned by
        Get-PSSessionConfiguration.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$PSSessionConfiguration = 'PowerShell.7',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Monitor\Monitor file count\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Error.Clear()

        #region Create log folder
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

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        try {
            if (-not ($MailTo = $file.MailTo)) {
                throw "Property 'MailTo' not found"
            }

            if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
                throw "Property 'MaxConcurrentJobs' not found"
            }
            try {
                $null = $MaxConcurrentJobs.ToInt16($null)
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
            }

            if (-not ($Tasks = $file.Tasks)) {
                throw "Property 'Tasks' not found"
            }

            foreach ($task in $Tasks) {
                if (-not $task.MaxFiles) {
                    throw "Property 'Tasks.MaxFiles' not found"
                }

                try {
                    $null = [int]$task.MaxFiles
                }
                catch {
                    throw "Property 'Tasks.MaxFiles' needs to be a number, the value '$($task.MaxFiles)' is not supported."
                }

                if (-not $task.ComputerName) {
                    throw "Property 'Tasks.ComputerName' not found"
                }

                if (-not $task.Path) {
                    throw "Property 'Tasks.Path' not found"
                }
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Add properties
        $Tasks = $Tasks | Select-Object -ExcludeProperty 'Job' -Property *,
        @{
            Name       = 'Job'
            Expression = {
                [PSCustomObject]@{
                    Result = $null
                    Error  = $null
                }
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $scriptBlock = {
            try {
                $task = $_

                #region Declare variables for parallel execution
                if (-not $MaxConcurrentJobs) {
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                }
                #endregion

                $invokeParams = @{
                    ComputerName      = $task.ComputerName
                    ArgumentList      = $task.Path, $task.MaxFiles
                    ConfigurationName = $PSSessionConfiguration
                    ErrorAction       = 'Stop'
                }

                $task.Job.Result = Invoke-Command @invokeParams -ScriptBlock {
                    Param (
                        [Parameter(Mandatory)]
                        [String]$Path,
                        [Parameter(Mandatory)]
                        [Int]$MaxFiles
                    )

                    if (-not (Test-Path -LiteralPath $Path -PathType 'Container')) {
                        throw "Path '$Path' not found"
                    }

                    $params = @{
                        LiteralPath = $Path
                        File        = $true
                        ErrorAction = 'Stop'
                    }
                    $fileCount = (Get-ChildItem @params | Measure-Object).Count

                    [PSCustomObject]@{
                        FileCount = $fileCount
                        IsTooMuch = $fileCount -gt $MaxFiles
                    }
                }
            }
            catch {
                $task.Job.Error = $_
                $Error.RemoveAt(0)
            }
        }

        #region Run code serial or parallel
        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{
                Process = $scriptBlock
            }
        }
        else {
            @{
                Parallel      = $scriptBlock
                ThrottleLimit = $MaxConcurrentJobs
            }
        }
        #endregion

        $Tasks | ForEach-Object @foreachParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        #region Get tasks with too many files
        $tasksWithTooManyFiles = $tasks | Where-Object {
            $_.Job.Result.IsTooMuch
        }
        #endregion

        #region Count total files
        $totalFileCount = (
            $tasksWithTooManyFiles.Job.Result |
            Measure-Object -Property 'FileCount' -Sum
        ).Sum + 0
        #endregion

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = '{0} files' -f $totalFileCount
            Priority  = 'High'
            Message   = @()
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($tasksWithTooManyFiles) {
            $htmlRows = $tasksWithTooManyFiles | ForEach-Object {
                $M = "Create HTML row for FileCount '$($_.Job.Result.FileCount)' ComputerName '$($_.ComputerName)' Path '$($_.Path)' MaxFiles '$($_.MaxFiles)'"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                @"
                <tr">
                    <td id="TxtLeft">$(
                        ConvertTo-HTMLlinkHC -Path $_.Path -Name $_.Path)</td>
                    <td id="TxtCentered">$($_.ComputerName)</td>
                    <td id="TxtCentered">$($_.Job.Result.FileCount)</td>
                    <td id="TxtCentered">$($_.MaxFiles)</td>
                </tr>
"@
            }

            $mailParams.Message += @"
                <style>
                #TxtLeft{
                    border: 1px solid Gray;
	                border-collapse:collapse;
	                text-align:left;
                }
                #TxtCentered {
	                text-align: center;
	                border: 1px solid Gray;
                }
                </style>
                <p>We found more files than indicated by '<b>MaxFiles</b>':</p>
                <table id="TxtLeft">
                <tr bgcolor="LightGrey" style="background:LightGrey;">
                <th id="TxtLeft">Path</th>
                <th id="TxtCentered" class="Centered">ComputerName</th>
                <th id="TxtCentered" class="Centered">FileCount</th>
                <th id="TxtCentered" class="Centered">MaxFiles</th>
                </tr>
                $htmlRows
                </table>
"@
        }

        #region Report errors
        $allErrors = @()

        if ($Error.Exception.Message) {
            $allErrors += $Error.Exception.Message
        }

        foreach (
            $task in
            $tasks | Where-Object { $_.Job.Error }
        ) {
            $allErrors += "Path '{0}' ComputerName '{1}' MaxFiles '{2}'<br>Error: {3}" -f
            $task.Path, $task.ComputerName, $task.MaxFiles, $task.Job.Error
        }

        if ($allErrors) {
            $mailParams.Subject = '{0}, {1} error{2}' -f
            $mailParams.Subject,
            $allErrors.Count,
            $(if ($allErrors.Count -ne 1) { 's' })

            $allErrors | ForEach-Object {
                Write-EventLog @EventErrorParams -Message "Error:`n`n- $_"
            }
            $mailParams.Message += $allErrors |
            ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors:'
        }
        #endregion

        if ($mailParams.Message) {
            Get-ScriptRuntimeHC -Stop
            Send-MailHC @mailParams
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}