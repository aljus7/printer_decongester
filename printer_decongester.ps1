$config = [PSCustomObject]@{
    AutomaticallyCloseKnownProcesses = $true
    AutomaticallyDenyUnknownProcesses = $true
    AlwaysPromptForPotentialPrinterDriverFix = $true
    AutomaticallyInstallHandle = $true
    AutomaticallyFixPrinterDriverIssues = $true
    DefaultPrinter = ""
    AssistedDriverReinstall = $true
    UsePnputilScanDevices = $true
    PrinterQueueCleanupRetryCount = 3
}

$printProcesses = @(
    "PrintIsolationHost",
    "PrintWorkflowUserSvc",
    "PrintWorkflowService",
    "PrintPipelineHost",
    "printfilterpipelinesvc"
)

$globalVars = [PSCustomObject]@{
    PrintersBeforeRemove = @(Get-Printer)
}

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
        if ($config.AutomaticallyInstallHandle -eq $true) {
            Write-Host "Automatically installing handle.exe" -ForegroundColor Yellow
            $response = "Y"
        } else {
            Write-Host "handle.exe not found — do you want to install it?" -ForegroundColor Red
            $response = Read-Host "Install handle.exe? (y/N)"
        }
        if ($response -match "^[Yy]$") {
            winget install --id Microsoft.Sysinternals.Handle -e --silent --accept-package-agreements --accept-source-agreements
            if ($?) {
                Write-Host "handle.exe installed successfully." -ForegroundColor Green
                # refresh paths withount restarting the shell
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
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

    if (@(Get-Printer).Count -eq 0) {
        Write-Host "No printers found." -ForegroundColor Red
        return
    }

    # Check if the default printer is set in the config and if not, find the printer with the most jobs to set as default 
    if ($config.AutomaticallyFixPrinterDriverIssues -and $config.DefaultPrinter -eq "") {
        Write-Host "No default printer set in config, attempting to find printer with most jobs to set as default for potential driver fixes."
        $config.DefaultPrinter = GetPrinterWithMostJobs
    }

    $closeAllKnown = $false;
    $handleAvailable = $false;
    if (-not (Get-Command handle.exe -ErrorAction SilentlyContinue)) {
        $handleAvailable = Install-Handle
    } else {
        $handleAvailable = $true
    }

    $retry = $config.PrinterQueueCleanupRetryCount
    if ($retry -lt 1) {
        $retry = 1
    }
    $can_break = $false
    while($retry -gt 0) {
        $filesPrint = @(Get-ChildItem -Path "C:\Windows\System32\spool\PRINTERS\" -Recurse)
        Stop-Service -Name Spooler -ErrorAction SilentlyContinue  
        if ($?) {
            if ($retry -eq 1 -and $handleAvailable) {
                Write-Host "Taking drastic measures, closing processes that own files!" -ForegroundColor Yellow
                if ($handleAvailable) {
                    foreach ($file in $filesPrint) {
                        $output = handle.exe -accepteula $file.FullName
                        $pidMatches = ($output | Select-String -Pattern "pid:\s+(\d+)").Matches
                        if ($pidMatches.Count -gt 0) {
                            foreach ($match in $pidMatches) {
                                $piid = $match.Groups[1].Value
                                # Get process name
                                $proc = Get-Process -Id $piid -ErrorAction SilentlyContinue
                                Write-Host "File: $($file.FullName) is being used by PID: $piid, Process: $proc.Name" -ForegroundColor Yellow

                                if ($printProcesses -contains $proc.Name -and $closeAllKnown -eq $false) {
                                    if ($config.AutomaticallyCloseKnownProcesses -eq $true) {
                                        Write-Host "Automatically closing known safe process $piid ($($proc.Name))" -ForegroundColor Yellow
                                        $closeAllKnown = $true
                                    } else {
                                        Write-Host "This is one of the print pipeline processes. Do you want to terminate all known safe processes?" -ForegroundColor Cyan
                                        $response = Read-Host "Terminate process $piid? (Y/N)"
                                        if ($response -match "^[Yy]$") {
                                            $closeAllKnown = $true
                                        }
                                    }
                                }
                                $response = "";
                                if ($closeAllKnown -eq $true -and $printProcesses -contains $proc.Name) {
                                    $response = "Y";
                                } 
                                if ($config.AutomaticallyDenyUnknownProcesses -eq $true -and $printProcesses -notcontains $proc.Name) {
                                    Write-Host "Automatically denying unknown process $piid ($($proc.Name))" -ForegroundColor Yellow
                                    $response = "N";
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
    if ($retry -eq 0 -or $config.AlwaysPromptForPotentialPrinterDriverFix -eq $true -or $config.AutomaticallyFixPrinterDriverIssues -eq $true) {
        Write-Host "Print a test page. And confirm if printing now works?"
        Write-Host "Found printer devices:"
        $printers = @(Get-Printer | Select-Object Name)
        $defaultPrinterIndex = -1
        for ($i = 0; $i -lt $printers.Count; $i++) {
            Write-Host "$($i): $($printers[$i].Name)"
            if ($config.AutomaticallyFixPrinterDriverIssues) {
                if ($config.DefaultPrinter -ne "" -and $printers[$i].Name -like "*$config.DefaultPrinter*") {
                    Write-Host "   (Default Printer to test)"
                    $defaultPrinterIndex = $i
                }
            }
        }
        $printerIndex = -1
        while ($true) {
            if ($defaultPrinterIndex -ge 0) {
                $printerIndex = $defaultPrinterIndex
                Write-Host "Using default printer: $($printers[$printerIndex].Name) for test page."
            } else {
                Write-Host "Enter the number of the printer you want to test:"
                $printerIndex = Read-Host "Printer number"
            }
            if ($printerIndex -gt $printers.Count-1 -or $printerIndex -lt 0) {
                Write-Host "Invalid printer number." -ForegroundColor Red
            } elseif ($defaultPrinterIndex -ge 0) {
                break
            } else {
                break
            }
        } 
        $printer = Get-CimInstance Win32_Printer -Filter "Name='$($printers[$printerIndex].Name)'"
        Invoke-CimMethod -InputObject $printer -MethodName PrintTestPage
        if ($config.AutomaticallyFixPrinterDriverIssues) {
            Write-Host "Automated page success validation in progress..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            $printingUserCheckWorks = waitForJobQueToEmpty $($printers[$printerIndex].Name)
        } else {
            $response = Read-Host "Did the test page print successfully? (Y/N)"
            if ($response -match "^[Yy]$") {
                $printingUserCheckWorks = $true
            } else {
                $printingUserCheckWorks = $false
            }
        }
        if (-not $printingUserCheckWorks -and $printerIndex -ge 0) {
            if ($config.AutomaticallyFixPrinterDriverIssues) {
                Write-Host "Attempting to fix printer driver issues automatically..." -ForegroundColor Yellow
                FixPrinterDriverIssues $($printers[$printerIndex].Name)
            } else {
            Write-Host "Printing still does not work. Do you want to attempt to fix printer driver issues?" -ForegroundColor Yellow
            $response = Read-Host "Attempt to fix printer driver issues? (Y/N)"
                if ($response -match "^[Yy]$") {
                    FixPrinterDriverIssues $($printers[$printerIndex].Name)
                }
            }
        }
    }

    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function GetPrinterWithMostJobs {
    $ret = ""
    $printerJobs = @(Get-Printer | Sort-Object JobCount -Descending | Select-Object Name, JobCount)
    if ($printerJobs.Count -gt 0 -and $printerJobs[0].JobCount -gt 0) {
        Write-Host "Setting default printer to $($printerJobs[0].Name) for test page."
        $ret = $printerJobs[0].Name
    }
    return $ret
}

function waitForJobQueToEmpty {
    param(
        [string]$printerName
    )
    $queueCount = @(Get-PrintJob -PrinterName $printerName -ErrorAction SilentlyContinue).Count
    $originalQueueCount = $queueCount
    Write-Host "Waiting for print jobs to complete..." -ForegroundColor Yellow
    $seconds = 45
    for ($i = 0; $i -lt $seconds; $i++) {
        Start-Sleep -Seconds 1
        Write-Host "`rProgress: $([int](((($originalQueueCount-$queueCount)+1)/($originalQueueCount+1))*100))%, Jobs remaining: $queueCount" -NoNewline
        $queueCount = @(Get-PrintJob -PrinterName $printerName -ErrorAction SilentlyContinue).Count
        if ($queueCount -eq 0) {
            Write-Host "`rProgress: 100%, Jobs remaining: 0" -NoNewline
            Write-Host "`nAll print jobs completed." -ForegroundColor Green
            return $true
        }
    }
    Write-Host "`nTimeout reached. Some print jobs may still be processing." -ForegroundColor Yellow
    return $false
}

function FixPrinterDriverIssues {
    param(
        [string]$printerName
    )
    $printer = Get-Printer -Name $printerName

    if ($printerName -eq "" -or $printerName -eq $null) {
        Write-Host "No printer name provided. Cannot fix driver issues." -ForegroundColor Red
        return
    }
    if ($printerName.Contains("PDF") -or $printerName.Contains("XPS") -or $printerName.Contains("Fax") -or $printerName.Contains("OneNote")) {
        Write-Host "Cannot fix driver issues for virtual printers." -ForegroundColor Red
        return
    }

    if ($printer.DriverName -like "*HP*" -or $printerName -like "*HP*") {
        DriverFix-HP $printerName
    }
}

function findReaddedPrinterDriver {
    param(
        [string]$printerName
    )
    $printers = @(Get-Printer | Select-Object Name)
    foreach ($printer in $printers) {
        if ($printer.Name -like "*$printerName*") {
            return $true
        }
        if ($printerName -like "*$printer.Name*") {
            return $true
        }
    }
    return $false
}

function waitForPrinterDriverReadd {
    param(
        [string]$printerName
    )
    $printersBefore = @(Get-Printer)
    for ($i = 45; $i -gt 0; $i--) {
        Write-Host "Waiting for printer driver to be re-added... ($i seconds remaining)" -NoNewline
        if ((findReaddedPrinterDriver $printerName) -and $printersBefore.Count -lt @(Get-Printer).Count) {
            Write-Host "`nPrinter driver re-added successfully." -ForegroundColor Green
            return $true
        }
        Start-Sleep -Seconds 1
    }
    Write-Host "`nPrinter driver did not re-add automatically!" -ForegroundColor Yellow
    return $false
}

function setPrintersBeforeRemove {
    $globalVars.PrintersBeforeRemove = @(Get-Printer)
}

function waitForPrinterRemove {
    for ($i = 60; $i -gt 0; $i--) {
        Write-Host "Waiting for printer driver to be removed... ($i seconds remaining)" -NoNewline
        if ($globalVars.PrintersBeforeRemove.Count -gt @(Get-Printer).Count) {
            Write-Host "`nPrinter driver removed successfully." -ForegroundColor Green
            return $true
        }
        Start-Sleep -Seconds 1
    }
    Write-Host "`nPrinter driver did not remove automatically!" -ForegroundColor Yellow
    return $false
}

function reAddPrinterDriver {
    param(
        [string]$printerName
    )
    if ($config.AssistedDriverReinstall) {
        if ($config.UsePnputilScanDevices) {
            Write-Host "Attempting to re-add printer driver automatically using pnputil..."
            pnputil /scan-devices
            if (waitForPrinterDriverReadd $printerName) {
                return
            }
            Write-Host "`nPrinter driver did not re-add automatically!" -ForegroundColor Yellow
        }
        Write-Host "Please re-add $printerName printer drivers manually." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        Start-Process "ms-settings:printers"
        if (waitForPrinterDriverReadd $printerName) {
            return
        }
    } else {
        Write-Host "Please re-add $printerName printer drivers." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Start-Process "ms-settings:printers"
        while ($true) {
            $response = Read-Host "Have you re-added the printer drivers? (Y/N)"
            if ($response -match "^[Yy]$") {
                return
            } else {
                Write-Host "Please re-add the $printerName printer drivers before continuing." -ForegroundColor Yellow
            }
        }
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
    setPrintersBeforeRemove 
    Remove-Printer -Name $printerName -ErrorAction SilentlyContinue
    waitForPrinterRemove
    get-appxpackage *HPPrinterDriver* -AllUsers | remove-appxpackage -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    if ($null -ne $printerDriverName) {
        Write-Host "Removing printer driver: $printerDriverName" -ForegroundColor Yellow
        if ($printerDriverName -like "*IPP*" -and $printerDriverName -like "*Microsoft*") {
            Write-Host "Detected IPP or Microsoft driver. Skipping removal." -ForegroundColor Yellow
        } else {
            Remove-PrinterDriver -Name $printerDriverName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2 
        }
    }
    reAddPrinterDriver $printerName
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