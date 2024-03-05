#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        MailTo            = 'bob@contoso.com'
        MaxConcurrentJobs = 1
        Tasks             = @(
            @{
                ComputerName = 'localhost'
                Path         = (New-Item 'TestDrive:/a' -ItemType Directory).FullName
                MaxFiles     = 2
            }
        )
    }

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $mailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MailTo',
                'MaxConcurrentJobs',
                'Tasks'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'MaxConcurrentJobs is not a number' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.MaxConcurrentJobs = 'a'

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' needs to be a number, the value 'a' is not supported*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'Tasks' {
                It '<_> not found' -ForEach @(
                    'ComputerName',
                    'Path'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and
                            ($Message -like "*$ImportFile*Property 'Tasks.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'MaxFiles is not a number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].MaxFiles = 'a'

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Tasks.MaxFiles' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
    It 'a Path in Tasks does not exist' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            @{
                ComputerName = 'localhost'
                Path         = 'TestDrive:\NotExisting'
                MaxFiles     = 2
            }
        )

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '0 files, 1 error') -and
            ($Message -like "*Errors:*Path*$($testNewInputFile.Tasks[0].Path)*ComputerName*$($testNewInputFile.Tasks[0].ComputerName)*MaxFiles*$($testNewInputFile.Tasks[0].MaxFiles)*Error: Path '$($testNewInputFile.Tasks[0].Path)' not found*")
        }
    }
    It 'PSSessionConfiguration is incorrect' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            @{
                ComputerName = 'localhost'
                Path         = 'TestDrive:\NotExisting'
                MaxFiles     = 2
            }
        )

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '0 files, 1 error') -and
            ($Message -like "*Errors:*Path*$($testNewInputFile.Tasks[0].Path)*ComputerName*$($testNewInputFile.Tasks[0].ComputerName)*MaxFiles*$($testNewInputFile.Tasks[0].MaxFiles)*Error: Path '$($testNewInputFile.Tasks[0].Path)' not found*")
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        1..5 | ForEach-Object {
            New-Item -Path "$($testInputFile.Tasks[0].Path)\$_.txt" -ItemType File -Force
        }

        $testInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        .$testScript @testParams
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '5 files') -and
            ($Message -like "*We found more files than indicated by '<b>MaxFiles</b>'*
            *Path*ComputerName*FileCount*MaxFiles*
            *$($testInputFile.Tasks[0].Path)*$($testInputFile.Tasks[0].ComputerName)*5*$($testInputFile.Tasks[0].MaxFiles)*")
        }
    }
}