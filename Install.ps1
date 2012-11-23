$modulePath = $ENV:userprofile + "\Documents\WindowsPowershell\Modules\SCCM"

if((Test-Path $modulePath) -ne $true) {
    New-Item $modulePath -ItemType directory > $null
}

try {
    Copy-Item -Path ".\SCCM.psd1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Advertisement.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Client.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Collection.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Computer.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Folder.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Package.psm1" -Destination $modulePath -Force
    Copy-Item -Path ".\SCCM_Formats.ps1xml" -Destination $modulePath -Force

    Write-Host ">>> SCCM module successfully installed to:" -Foreground green
    Write-Host ">>> `t$modulePath" -Foreground green
    Write-Host ">>>"
    Write-Host ">>> Make sure that $modulePath is part of your PSModulePath"
    Write-Host ">>> then try running:"
    Write-Host ">>>"
    Write-Host ">>> `tImport-Module SCCM"
    Write-Host ">>> `tGet-Help SCCM"
    Write-Host ">>>"
} catch {
    Write-Host ">>> There was a problem when trying to install to:" -Foreground red
    Write-Host ">>> `t$modulePath" -Foreground red
    Write-Host ">>>"
    Write-Host ">>> You might need to perform the installation manually."
    Write-Host '>>>'
}