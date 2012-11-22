<#
.SYNOPSIS
Creates a new computer account in SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and a MAC address in order to create a new 
account in SCCM.

.PARAMETER siteProvider
The name of the site provider for the site where the new computer account is to be created.

.PARAMETER siteCode
The 3-character site code for the site where the new computer account is to be created.

.PARAMETER computerName
Name of the computer to be created.  Name should not be longer than 15 characters.

.PARAMETER macAddress
MAC address of the computer account to be created in the format 00:00:00:00:00:00.

.EXAMPLE
New-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT -computerName MYCOMPUTER -macAddress "00:00:00:00:00:00"
#>
Function New-SCCMComputer {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(1,15)]
        [ValidateNotNullOrEmpty()]
        [string]$computerName,
        [parameter(Mandatory=$true, Position=1)]
        [ValidatePattern('^([0-9A-F]{2}[:-]){5}([0-9A-F]{2})$')]
        [string]$macAddress
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $site = [WMIClass]("\\$siteProvider\ROOT\sms\site_" + $siteCode + ":SMS_Site")

    $methodParameters = $site.GetMethodParameters("ImportMachineEntry")
    $methodParameters.MACAddress = $macAddress
    $methodParameters.NetbiosName = $computerName
    $methodParameters.OverwriteExistingRecord = $false

    $computerCreationResult = $site.InvokeMethod("ImportMachineEntry", $methodParameters, $null)

    if($computerCreationResult.MachineExists -eq $true) {
        Throw "Computer already exists with name $computerName or MAC $macAddress"
    } elseif($computerCreationResult.ReturnValue -eq 0) {
        return Get-SCCMComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $computerCreationResult.ResourceID
    } else {
        Throw "Computer account creation failed for $computerName"
    }
}

<#
.SYNOPSIS
Removes a computer account in SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID and will attempt to delete it from the SCCM site.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site to be queried.

.PARAMETER resourceId
Resource ID of the computer that needs to be deleted.

.EXAMPLE
Remove-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT -resourceId 1293

Description
-----------
Removes the computer with resource ID 1293 from the specified site.
#>
Function Remove-SCCMComputer {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateScript( { $_ -gt 0 } )]
        [parameter(Mandatory=$true, Position=0)][int]$resourceId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $computer =  Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System" | where { $_.ResourceID -eq $resourceId }
    if($computer) {
        $computer.Delete() | Out-Null
    } else {
        Throw "Unable to retrieve computer with resource ID $resourceId"
    }
}

<#
.SYNOPSIS
Returns computers from SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name or resource ID and will attempt to retrieve an object for the computer.  
When invoked without specifying a computer name or resource ID, this function returns a list of all computers found on the specified SCCM site.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site where the computer exists.

.PARAMETER computerName
The name of the computer to be retrieved.

.PARAMETER resourceId
The resource ID of the computer to be retrieved.

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT -computerName MYCOMPUTER

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT MYCOMPUTER

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT -resourceId 1111

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT
#>
Function Get-SCCMComputer {
    [CmdletBinding(DefaultParametersetName="default")]
    param (
        [parameter(ParameterSetName="name")]
        [parameter(ParameterSetName="default")]
        [parameter(ParameterSetName="id")]
        [string]$siteProvider,
        [parameter(ParameterSetName="name")]
        [parameter(ParameterSetName="default")]
        [parameter(ParameterSetName="id")]
        [string]$siteCode,
        [parameter(Position=0, ValueFromPipeline=$true)]
        [parameter(ParameterSetName="name")]
        [ValidateLength(1,15)]
        [string]$computerName,
        [parameter(ParameterSetName="id")]
        [ValidateScript( { $_ -gt 0 } )]
        [int]$resourceId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($computerName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System" | where { $_.Name -like $computerName }
    } elseif($resourceId) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System" | where { $_.ResourceID -eq $resourceId }
    } else {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System"
    }
}

<#
.SYNOPSIS
Returns the advertisement status for a specific advertisement for a specific computer.

.DESCRIPTION
Takes in information about a specific site, along with an advertisement id and a computer resource ID and returns the status of that advertisement for that computer.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character code for the site where the computer exists.

.PARAMETER advertisementId
The ID of the advertisement.

.PARAMETER resourceId
The resource ID of the computer.
#>
Function Get-SCCMAdvertisementStatusForComputer {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$advertisementId,
        [parameter(Mandatory=$true, Position=1)][int]$resourceId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_ClientAdvertisementStatus WHERE AdvertisementID='$advertisementId' AND ResourceID='$resourceId'"
}

<#
.SYNOPSIS
Sets variables on an SCCM computer record.

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID and an array of variables, and assigns those variables to the computer.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
Resource ID of the target computer.

.PARAMETER variableList
An array of variables to be assigned to the computer.  If you pass it an empty array, it will clear all variables on the computer.

.LINK
http://msdn.microsoft.com/en-us/library/cc143033.aspx
#>
Function Set-SCCMComputerVariables {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][int]$resourceId,
        [parameter(Mandatory=$true, Position=1)][ValidateNotNull()]$variableList
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $computer = Get-SCCMComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    if(!$computer) {
        Throw "Unable to retrieve computer with resource ID $resourceId"
    }

    $computerSettings = Get-SCCMMachineSettings -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    if(!$computerSettings) {
        $computerSettings = New-SCCMMachineSettings -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    }
    $computerSettings.MachineVariables = $variableList
    Save-SCCMMachineSettings $computerSettings
}

<#
.SYNOPSIS
Returns an array of all computer variables for a specific computer.

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID and returns a list of all variables associated with that computer.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
Resource ID of the target computer.

.LINK
http://msdn.microsoft.com/en-us/library/cc143033.aspx
#>
Function Get-SCCMComputerVariables {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][int]$resourceId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $computer = Get-SCCMComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    if(!$computer) {
        Throw "Unable to retrive computer with resource ID $resourceId"
    }

    $computerSettings = Get-SCCMMachineSettings -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    if($computerSettings) {
        return $computerSettings.MachineVariables
    }
}

<#
.SYNOPSIS
Creates a new computer variable.

.DESCRIPTION
Creates a new computer variable.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER variableName
Name of the variable to be created.

.PARAMETER variableValue
Value to be assigned to the variable.

.PARAMETER isMasked
A masked variable has its text display replaced with asterisks.

.LINK
http://msdn.microsoft.com/en-us/library/cc143033.aspx
#>
Function New-SCCMComputerVariable {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true)][string]$variableName,
        [parameter(Mandatory=$true)][string]$variableValue,
        [parameter(Mandatory=$true)][bool]$isMasked
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $machineVariable = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_MachineVariable")).CreateInstance()
    if($machineVariable) {
        $machineVariable.IsMasked = $isMasked
        $machineVariable.Name = $variableName
        $machineVariable.Value = $variableValue
    }
    return $machineVariable
}

<#
.SYNOPSIS
Returns the machine settings for a computer.

.DESCRIPTION
Returns the machine settings for a computer.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
Resource ID of the computer whose settings are to be returned.

.LINK
http://msdn.microsoft.com/en-us/library/cc145852.aspx
#>
Function Get-SCCMMachineSettings {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [int]$resourceId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $machineSettings = Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_MachineSettings" | where {$_.ResourceID -eq $resourceId}
    if($machineSettings) {
        $machineSettings.Get() | Out-Null
    }
    return $machineSettings
} 

<#
.SYNOPSIS
Creates a new object for a computer's machine settings.

.DESCRIPTION
Creates a new object for a computer's machine settings.  These settings are not saved to the database during creation in this method, and must
be explicitly saved once you are finished configuring them.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
The resource ID of the computer whose settings are being created.

.LINK
http://msdn.microsoft.com/en-us/library/cc145852.aspx
#>
Function New-SCCMMachineSettings {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)]$resourceId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $computer = Get-SCCMComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    if(!$computer) {
        Throw "Unable to retrive computer with resource ID $resourceId"
    }

    $machineSettings = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_MachineSettings")).CreateInstance()
    if($machineSettings) {
        $machineSettings.ResourceID = $computer.ResourceID
        $machineSettings.SourceSite = $siteCode
        $machineSettings.LocaleID = $computer.LocaleID
    }
    return $machineSettings
}

<#
.SYNOPSIS
Saves an object containing a computer's machine settings back into the database.

.DESCRIPTION
Saves an object containing a computer's machine settings back into the database.

.PARAMETER machineSettings
The object to be saved back into the database.

.LINK
http://msdn.microsoft.com/en-us/library/cc145852.aspx
#>
Function Save-SCCMMachineSettings {
    [CmdletBinding()]
    param (        
        [parameter(Mandatory=$true)]$machineSettings
    )

    $machineSettings.Put() | Out-Null
}