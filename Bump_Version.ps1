param (
    [parameter(Mandatory=$true, ParameterSetName="release")][switch]$release,
    [parameter(Mandatory=$true, ParameterSetName="feature")][switch]$feature,
    [parameter(Mandatory=$true, ParameterSetName="fix")][switch]$fix
)

$moduleFile = ".\SCCM.psd1"
$searchPattern = "^[\s]*ModuleVersion[\s]*=[\s]*`'([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})`'"

$match = Select-String -Path $moduleFile -Pattern $searchPattern
if($match) {
    # We're only interested in the first match
    $firstMatch = $($match.Matches)[0]
    $currentVersionString = $($firstMatch.Value.Split("`'"))[1]
    $currentVersionParts = $currentVersionString.Split(".")

    $versionNumber = [int]$currentVersionParts[0]
    $featureNumber = [int]$currentVersionParts[1]
    $fixNumber = [int]$currentVersionParts[2]

    if($release) {
        $versionNumber++
        $featureNumber = 0
        $fixNumber = 0
    } elseif($feature) {
        $featureNumber++
    } else {
        $fixNumber++
    }

    Write-Host "Previous version:`t$currentVersionString"
    Write-Host "Bumped version:`t`t$versionNumber.$featureNumber.$fixNumber"

    $moduleFileContents = Get-Content $moduleFile
    $newVersionString = "ModuleVersion = `'" + $versionNumber + "." + $featureNumber + "." + $fixNumber + "`'"
    $moduleFileContents -Replace $searchPattern, $newVersionString | Out-File $moduleFile
}



