$modulePath = $ENV:userprofile + "\Documents\WindowsPowershell\Modules\SCCM"

if((Test-Path $modulePath) -ne $true) {
    New-Item $modulePath -ItemType directory > $null
}

try {
    Copy-Item -Path ".\SCCM.psd1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Formats.ps1xml" -Destination $modulePath -Force

    Write-Host ">>> SCCM module successfully installed to $modulePath" -Foreground green
    Write-Host ">>> Make sure that $modulePath is part of your PSModulePath then try running Import-Module SCCM followed by Get-Help SCCM" -Foreground Yellow
} catch {
    Write-Host "SCCM module failed to install to $modulePath" -Foreground red 
}