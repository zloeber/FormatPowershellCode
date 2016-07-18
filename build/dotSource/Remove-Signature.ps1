#requires -Version 2
function Script:Remove-Signature
{
    [cmdletbinding()]

    Param(
        [Parameter(Mandatory = $False,Position = 0,ValueFromPipeline = $True,ValueFromPipelineByPropertyName = $True)]
        [Alias('Path')]
        [system.io.fileinfo[]]$FilePath
    )

    Begin{
        Push-Location -Path $env:USERPROFILE
    }

    Process{
        $FilePath |
        ForEach-Object -Process {
            $Item = $_
			
            If($Item.Extension -match '\.ps1|\.psm1|\.psd1|\.ps1xml')
            {
                Try
                {
                    $Content = Get-Content -Path $Item.FullName -ErrorAction Stop
    
                    $StringBuilder = New-Object -TypeName System.Text.StringBuilder -ErrorAction Stop
    
                    Foreach($Line in $Content)
                    {
                        If($Line -match '^# SIG # Begin signature block|^<!-- SIG # Begin signature block -->')
                        {
                            Break
                        }
                        Else
                        {
                            $null = $StringBuilder.AppendLine($Line)
                        }
                    }
    
                    Set-Content -Path $Item.FullName -Value $StringBuilder.ToString()
                }
                Catch
                {
                    Write-Error -Message $_.Exception.Message
                }
            }
        }
    }

    End{
        Pop-Location
    }
}