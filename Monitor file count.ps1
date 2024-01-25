#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, Toolbox.Remoting

<#
    .SYNOPSIS
        Send an e-mail when there are more than x files in a folder.

    .DESCRIPTION
        Count all the files in a folder and send an e-mail when there are more
        than x files in the specified folder.

    .PARAMETER ImportFile
        Contains all the parameters for the script.

    .PARAMETER MailTo
        E-mail addresses of where to send the summary e-mail
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
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
                    Object = $null
                    Result = $null
                    Errors = @()
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
            Param (
                [Parameter(Mandatory)]
                [String]$Path,
                [Parameter(Mandatory)]
                [Int]$MaxFiles
            )

            if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
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

        #region Get file count
        foreach ($task in $Tasks) {
            try {
                Write-Verbose "Start job on '$($task.ComputerName)' for path '$($task.Path)'"

                # & $scriptBlock -Path $task.Path -MaxFiles $task.MaxFiles

                #region Start job
                $invokeParams = @{
                    ScriptBlock  = $scriptBlock
                    Session      = New-PSSessionHC -ComputerName $task.ComputerName
                    ArgumentList = $task.Path, $task.MaxFiles
                    AsJob        = $true
                }
                $task.Job.Object = Invoke-Command @invokeParams
                #endregion

                #region Wait for max running jobs
                if ($Tasks.Job.Object) {
                    $waitParams = @{
                        Name       = $Tasks.Job.Object | Where-Object { $_ }
                        MaxThreads = $MaxConcurrentJobs
                    }
                    Wait-MaxRunningJobsHC @waitParams
                }
                #endregion
            }
            catch {
                $task.Job.Errors += $_
                $Error.RemoveAt(0)
            }
        }
        #endregion

        #region Wait for jobs to finish
        if ($Tasks.Job.Object) {
            Write-Verbose 'Wait for all jobs to finish'
            $Tasks.Job.Object | Where-Object { $_ } | Wait-Job
        }

        $M = 'All jobs finished'
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Get job results and job errors
        foreach (
            $task in
            $Tasks | Where-Object { $_.Job.Object }
        ) {
            Write-Verbose "Get job result for ComputerName '$($task.ComputerName)' Path '$($task.Path)' MaxFiles '$($task.MaxFiles)'"

            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $task.Job.Result = $task.Job.Object | Receive-Job @receiveParams

            Write-Verbose "Job result FileCount '$($task.Job.Result.FileCount)' IsTooMuch '$($task.Job.Result.IsTooMuch)'"

            foreach ($e in $jobErrors) {
                Write-Warning "Failed with job error: $($e.ToString())"

                $task.Job.Errors += $e.ToString()
                $Error.Remove($e)
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
            $tasks | Where-Object { $_.Job.Errors }
        ) {
            foreach ($e in $task.Job.Errors) {
                $allErrors += "Path '{0}' ComputerName '{1}' MaxFiles '{2}'<br>Error: {3}" -f
                $task.Path, $task.ComputerName, $task.MaxFiles, $e
            }
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