# Check if the script is running with administrator rights
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ([int](Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $arguments = "& '" + $myinvocation.mycommand.definition + "'"
        Start-Process powershell -Verb runAs -ArgumentList $arguments
        exit
    }
}

Write-Host "SHELL WILL CLOSE AUTOMATICALLY!!" -ForegroundColor Yellow
$closeAllKnown = $false;

$printProcesses = @(
    "PrintIsolationHost",
    "PrintWorkflowUserSvc",
    "PrintWorkflowService",
    "PrintPipelineHost",
    "printfilterpipelinesvc"
)

$handleAvailable = $false;
if (-not (Get-Command handle.exe -ErrorAction SilentlyContinue)) {
    Write-Host "handle.exe not found — skipping handle-based detection." -ForegroundColor Red
} else {
    $handleAvailable = $true
}

$retry = 3
while($retry -gt 0) {
    $filesPrint = Get-ChildItem -Path "C:\Windows\System32\spool\PRINTERS\" -Recurse
    $can_break = $false
    Stop-Service -Name Spooler -ErrorAction SilentlyContinue  
    if ($?) {
        if ($retry -eq 2 -and $handleAvailable) {
            Write-Host "Taking drastic measures, closing processes that own files!" -ForegroundColor Yellow
            if ($handleAvailable) {
                foreach ($file in $filesPrint) {
                    $output = handle.exe -accepteula $file.FullName
                    $pidMatches = ($output | Select-String -Pattern "pid:\s+(\d+)").Matches
                    if ($pidMatches.Count -gt 0) {
                        $piid = $pidMatches[0].Groups[1].Value
                        # Get process name
                        $proc = Get-Process -Id $piid -ErrorAction SilentlyContinue
                        Write-Host "File: $($file.FullName) is being used by PID: $piid, Process: $proc.Name" -ForegroundColor Yellow

                        if ($printProcesses -contains $proc.Name -and $closeAllKnown -eq $false) {
                            Write-Host "This is one of the print pipeline processes\n Do you want to terminate all knows safe processes?" -ForegroundColor Cyan
                            $response = Read-Host "Terminate process $piid? (Y/N)"
                            if ($response -match "^[Yy]$") {
                                $closeAllKnown = $true
                            }
                        }
                        $response = "";
                        if ($closeAllKnown -eq $true -and $printProcesses -contains $proc.Name) {
                            $response = "Y";
                        }
                        if ($response -eq "") {
                            $response = Read-Host "Terminate process $piid? (Y/N)"
                        }
                        if ($response -match "^[Yy]$") {
                            Stop-Process -Id $piid -Force
                            Write-Host "Terminated Process ID: $piid" -ForegroundColor Orange
                        } else {
                            Write-Host "Skipped terminating PID: $piid" -ForegroundColor Cyan
                        }
                    }
                    else {
                        Write-Host "File: $($file.FullName) is not used by any process" -ForegroundColor Yellow
                    }
                }
            }
        }
        Start-Sleep -Seconds 5
        Remove-Item -Path "C:\Windows\System32\spool\PRINTERS\*" -Force
        if ($?) {
            Write-Host "Sucess, cleared spool/PRINTERS folder!" -ForegroundColor Green
            $can_break = $true
            Start-Sleep -Seconds 2
        } else {
            Write-Host "Unsucessful, cleared spool/PRINTERS folder!" -ForegroundColor Red
        }
    }
    Start-Service -Name Spooler -ErrorAction SilentlyContinue 
    if ($?) {
        Write-Host "Restarted spooler service" -ForegroundColor Green
        Start-Sleep -Seconds 3
    }
    else {
        Write-Host "Unsucessful, restarted spooler service" -ForegroundColor Red
        $can_break = $false
    }
    if ($can_break -eq $true) {
        break
    }
    Start-Sleep -Seconds 3
    $retry--
 }

 if ($can_break -eq $true) {
    Write-Host "All done, printers should work again now!" -ForegroundColor Green
 }

 if ($retry -eq 0) {
    Write-Host "All retrys where unsucessfull :(" -ForegroundColor Red
}

Start-Sleep -Seconds 3
