function Grant-AdminRights {
    # Check if the script is running with administrator rights
    if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        if ([int](Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
            $arguments = "& '" + $myinvocation.mycommand.definition + "'"
            Start-Process powershell `
                -Verb RunAs `
                -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
            exit
        }
    }
}

function Install-Handle {
    Write-Host "handle.exe not found — do you want to install it?" -ForegroundColor Red
        $response = Read-Host "Install handle.exe? (y/N)"
        if ($response -match "^[Yy]$") {
            winget install --id Microsoft.Sysinternals.Handle -e
            if ($?) {
                Write-Host "handle.exe installed successfully." -ForegroundColor Green
                # refresh paths withount restarting the shell
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment\]::GetEnvironmentVariable("Path","User")
                if (Get-Command handle.exe -ErrorAction SilentlyContinue) {
                    return $true
                } else {
                    Write-Host "handle.exe still not found, try restarting the shell." -ForegroundColor Red
                }
            } else {
                Write-Host "Failed to install handle.exe. Please install it manually." -ForegroundColor Red
            }
        }
        return $false
}

function Main {
    Grant-AdminRights
    Write-Host "DONT CANCEL THE PROCESS!!" -ForegroundColor Yellow
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
        $handleAvailable = Install-Handle
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
                                Write-Host "This is one of the print pipeline processes. Do you want to terminate all known safe processes?" -ForegroundColor Cyan
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
                                Write-Host "Terminated Process ID: $piid" -ForegroundColor Yellow
                            } else {
                                Write-Host "Skipped terminating PID: $piid" -ForegroundColor Cyan
                            }
                        }
                        else {
                            Write-Host "File: $($file.FullName) is not used by any process" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "handle.exe not found — skipping handle-based detection." -ForegroundColor Red
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

    $printingUserCheckWorks = $false;
    if ($retry -gt 1) {
        Write-Host "Print a test page. And confirm if printing now works?"
        Write-Host "Found printer devices:"
        $printers = Get-Printer | Select-Object Name
        for ($i = 0; $i -lt $printers.Count; $i++) {
            Write-Host "$($i): $($printers[$i].Name)"
        }
        $printerIndex = -1
        while ($true) {
            Write-Host "Enter the number of the printer you want to test:"
            $printerIndex = Read-Host "Printer number"
            if ($printerIndex -gt $printers.Count-1 -or $printerIndex -lt 0) {
                Write-Host "Invalid printer number." -ForegroundColor Red
            } else {
                break
            }
        } 
        $printer = Get-CimInstance Win32_Printer -Filter "Name='$($printers[$printerIndex].Name)'"
        Invoke-CimMethod -InputObject $printer -MethodName PrintTestPage
        $response = Read-Host "Did the test page print successfully? (Y/N)"
        if ($response -match "^[Yy]$") {
            $printingUserCheckWorks = $true
        } else {
            $printingUserCheckWorks = $false
        }

        if (-not $printingUserCheckWorks -and $printerIndex -ge 0) {
        Write-Host "Printing still does not work. Do you want to attempt to fix printer driver issues?" -ForegroundColor Yellow
        $response = Read-Host "Attempt to fix printer driver issues? (Y/N)"
            if ($response -match "^[Yy]$") {
                FixPrinterDriverIssues $($printers[$printerIndex].Name)
            }
        }
    }

    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function FixPrinterDriverIssues {
    param(
        [string]$printerName
    )
    if ($printerName -eq "" -or $printerName -eq $null) {
        Write-Host "No printer name provided. Cannot fix driver issues." -ForegroundColor Red
        return
    }
    if ($printerName.Contains("PDF") -or $printerName.Contains("XPS") -or $printerName.Contains("Fax") -or $printerName.Contains("OneNote")) {
        Write-Host "Cannot fix driver issues for virtual printers." -ForegroundColor Red
        return
    }

    if ($printerName.contains("HP") -or $printerName.contains("hp")) {
        DriverFix-HP $printerName
    }
}

# HP Printer Driver Fix Function
function DriverFix-HP {
    param(
        [string]$printerName
    )
    Write-Host "Detected HP printer. Attempting to fix driver issues by reinstalling drivers." -ForegroundColor Yellow
    $printerDriverName = (Get-Printer |
        Where-Object Name -eq $printerName
        ).DriverName    
    Remove-Printer -Name $printerName -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    get-appxpackage *HPPrinterDriver* -AllUsers | remove-appxpackage -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    if ($null -ne $printerDriverName) {
        Write-Host "Removing printer driver: $printerDriverName" -ForegroundColor Yellow
        if ($printerDriverName.contains("IPP") -or $printerDriverName.contains("Microsoft")) {
            Write-Host "Detected IPP or Microsoft driver. Skipping removal." -ForegroundColor Yellow
        } else {
            Remove-PrinterDriver -Name $printerDriverName -ErrorAction SilentlyContinue -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2 
        }
    }
    Write-Host "Re adding HP printer drivers." -ForegroundColor Yellow
    Start-Process "ms-settings:printers"
    while ($true) {
        $response = Read-Host "Have you re-added the HP printer drivers? (Y/N)"
        if ($response -match "^[Yy]$") {
            break
        } else {
            Write-Host "Please re-add the HP printer drivers before continuing." -ForegroundColor Yellow
        }
    }
}

# calling main function
try {
    Main
}
catch {
    Write-Host "ERROR:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
     
    Write-Host ""
    Write-Host "Press any key to exit..."
    [void][Console\]::ReadKey($true)
}