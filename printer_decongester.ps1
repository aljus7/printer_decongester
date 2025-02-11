# Check if the script is running with administrator rights
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ([int](Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $arguments = "& '" + $myinvocation.mycommand.definition + "'"
        Start-Process powershell -Verb runAs -ArgumentList $arguments
        exit
    }
}

Write-Host "SHELL WILL CLOSE AUTOMATICALLY!!" -ForegroundColor Yellow

$retry = 3
while($retry -gt 0) {
    $can_break = $false
    Stop-Service -Name Spooler -ErrorAction SilentlyContinue  
    if ($?) {
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
        Write-Host "Sucess, restarted spooler service" -ForegroundColor Green
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

 if ($retry -eq 0) {
    Write-Host "All retrys where unsucessfull :(" -ForegroundColor Red
}
