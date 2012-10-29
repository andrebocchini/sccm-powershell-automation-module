$modulePath = $ENV:userprofile + "\Documents\WindowsPowershell\Modules\SCCM"

if((Test-Path $modulePath) -ne $true) {
    New-Item $modulePath -ItemType directory > $null
}
Copy-Item -Path ".\SCCM.psm1" -Destination $modulePath -Force

if(Test-Path "$modulePath\SCCM.psm1") {
    Write-Host "SCCM module successfully installed to $modulePath" -Foreground green
    Write-Host "Try running Import-Module SCCM followed by Get-Help SCCM" -Foreground green
} else {
    Write-Host "SCCM module failed to install to $modulePath" -Foreground red 
}