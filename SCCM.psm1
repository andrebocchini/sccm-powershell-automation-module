<#
.SYNOPSIS
Command line interface for an assortment of SCCM 2007 operations.

.DESCRIPTION
The functions in this module provide a command line and scripting interface for automating management of SCCM 2007 environments.

.NOTES  
File Name  : SCCM.ps1  
Author     : Andre Bocchini <andrebocchini@gmail.com>  
Requires   : PowerShell V2

.LINK
https://github.com/andrebocchini/SCCM-Powershell-Automation-Module
#>

#Requires -version 2

<#
.SYNOPSIS
Creates a new computer account in SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and a MAC address in order to create a new 
account in SCCM.  This function will not overwrite existing computers with a matching name or MAC.
#>
Function New-SCCMComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$macAddress
    )

    $site = [WMIClass]("\\$siteServer\ROOT\sms\site_" + $siteCode + ":SMS_Site")

    $methodParameters = $site.psbase.GetMethodParameters("ImportMachineEntry")
    $methodParameters.MACAddress = $macAddress
    $methodParameters.NetbiosName = $computerName
    $methodParameters.OverwriteExistingRecord = $false

    return $site.psbase.InvokeMethod("ImportMachineEntry", $methodParameters, $null)
}

<#
.SYNOPSIS
Removes a computer account in SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and will attempt to delete it from the SCCM server.
#>
Function Remove-SCCMComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    return $computer.psbase.Delete()
}

<#
.SYNOPSIS
Returns a computer object from SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and will attempt to retrieve a WMI object for the computer.  This function
intentionally ignores obsolete computers.

When invoked without specifying a computer name, this function returns a list of all computers found on the specified SCCM site.
#>
Function Get-SCCMComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [string]$computerName
    )

    if($computerName) {
        return Get-WMIObject -ComputerName $siteServer -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System" | where { ($_.Name -eq $computerName) -and ($_.Obsolete -ne 1) }
    } else {
        return Get-WMIObject -ComputerName $siteServer -Namespace "root\sms\site_$siteCode" -Query "select * from SMS_R_System"
    }
}

<#
.SYNOPSIS
Adds a computer to a collection in SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name, and a collection name and direct membership rule for the computer in
the specified collection.
#>
Function Add-SCCMComputerToCollection {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$collectionName
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    $collectionRule = [WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_CollectionRuleDirect")
    $collectionRule.Properties["ResourceClassName"].Value = "SMS_R_System"
    $collectionRule.Properties["ResourceID"].Value = $computer.Properties["ResourceID"].Value

    $collection = Get-WmiObject -ComputerName $siteServer -Namespace "root\sms\site_$siteCode" -Class "SMS_Collection" | where { $_.Name -eq $collectionName }
    $collectionId = $collection.collectionId
    $addToCollectionParameters = $collection.psbase.GetmethodParameters("AddMembershipRule")
    $addToCollectionParameters.collectionRule = $collectionRule  

    return $collection.psbase.InvokeMethod("AddMembershipRule", $addToCollectionParameters, $null)
}

<#
.SYNOPSIS
Removes a computer from a specific collection

.DESCRIPTION
Takes in information about a specific site, along with a computer name, and a collection name and removes the computer from the collection.
#>
Function Remove-SCCMComputerFromCollection {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$collectionName
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    $collection = Get-SCCMCollection $siteServer $siteCode $collectionName
    $collectionRule = ([WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_CollectionRuleDirect")).CreateInstance()
    $collectionRule.ResourceID = $computer.ResourceID
    
    return $collection.DeleteMembershipRule($collectionRule)
}

<#
.SYNOPSIS
Retrieves SCCM collection objects from the site server.

.DESCRIPTION
Takes in information about a specific site and a collection name and returns an object representing that collection.  If no computer name is specified, it returns all collections found on the site server.

.PARAMETER siteServer
The name of the site server to be queried.

.PARAMETER siteCode
The 3-character site code for the site to be queried.

.PARAMETER collectionName
Optional parameter.  If specified, the function returns an object representing the collection specified.  If absent, the function returns all collections for the site.

.EXAMPLE
Get-SCCMCollection -siteServer MYSITESERVER -siteCode SIT -collectionName MYCOLLECTION

Description
-----------
Retrieve the collection named MYCOLLECTION from site SIT on MYSITESERVER

.EXAMPLE
Get-SCCMCollection -siteServer MYSITESERVER -siteCode SIT

Description
-----------
Retrieve all collections from site SIT on MYSITESERVER

.EXAMPLE
Get-SCCMCollection -siteServer MYSITESERVER -siteCode SIT | Select-Object Name,CollectionID

Description
-----------
Retrieve all collections from site SIT on MYSITESERVER and filter out only their names and IDs
#>
Function Get-SCCMCollection {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [string]$collectionName
    )

    if($collectionName) {
        return Get-WMIObject -Computer $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Collection" | Where { $_.Name -eq $collectionName }  
    } else {
        return Get-WMIObject -Computer $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Collection"
    }
}

<#
.SYNOPSIS
Returns a list of collection IDs that a computer belongs to.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and returns an array containing the collection IDs that the computer belongs to.
#>
Function Get-SCCMCollectionsForComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName
    )

    # First we find all membership associations that match the colection ID of the computer in question
    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    $computerCollectionIds = @()
    $collectionMembers = Get-WMIObject -Computer $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_CollectionMember_a Where ResourceID= $($computer.ResourceID)"
    foreach($collectionMember in $collectionMembers) {
        $computerCollectionIds += $collectionMember.CollectionID
    }

    # Now that we have a list of collection IDs, we want to retrieve and return some rich collection objects
    $allServerCollections = Get-SCCMCollection $siteServer $siteCode
    $computerCollections = @()
    foreach($collectionId in $computerCollectionIds) {
        foreach($collection in $allServerCollections) {
            if($collection.CollectionID -eq $collectionId) {
                $computerCollections += $collection
            }
        }
    }

    return $computerCollections
}

<#
.SYNOPSIS
Returns a list of advertisements assigned to a collection.

.DESCRIPTION
Takes in information about a specific site, along with a collection ID and returns all advertisements assigned to that collection.
#>
Function Get-SCCMAdvertisementsForCollection {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$collectionId
    )

    return Get-WMIObject -Computer $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Advertisement WHERE CollectionID='$collectionId'"
}

<#
.SYNOPSIS
Returns a list of advertisements assigned to a computer.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and returns all advertisements assigned to that computer.
#>
Function Get-SCCMAdvertisementsForComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName
    )

    $collections = Get-SCCMCollectionsForComputer $siteServer $siteCode $computerName
    $computerAdvertisements = @()
    foreach($collection in $collections) {
        $advertisements = Get-SCCMAdvertisementsForCollection $siteServer $siteCode $collection.CollectionID
        foreach($advertisement in $advertisements) {
            $computerAdvertisements += $advertisement
        }
    }
    
    return $computerAdvertisements
}

<#
.SYNOPSIS
Returns the advertisement status for a specific advertisement for a specific computer.

.DESCRIPTION
Takes in information about a specific site, along with an advertisement id and a computer name and returns the status of that advertisement for that computer.
#>
Function Get-SCCMAdvertisementStatusForComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$advertisementId,
        [parameter(Mandatory=$true)][string]$computerName
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    return Get-WMIObject -Computer $siteServer -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_ClientAdvertisementStatus WHERE AdvertisementID='$advertisementId' AND ResourceID='$($computer.ResourceID)'"
}

<#
.SYNOPSIS
Settings a variable on an SCCM computer record.

.DESCRIPTION
Takes in information about a specific site, along with a computer name, variable name and value and attempts to set the variable on an SCCM computer record.
If the variable already exists, it will overwrite be overwritten.
#>
Function Set-SCCMComputerVariable {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$variableName,
        [parameter(Mandatory=$true)][string]$variableValue,
        [parameter(Mandatory=$true)][bool]$isMasked
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    $computerSettings = Get-WMIObject -computername $siteServer -namespace "root\sms\site_$siteCode" -class "SMS_MachineSettings" | where {$_.ResourceID -eq $computer.ResourceID}

    # If the computer has never held any variables, computerSettings will be null, so we have to create a bunch of objects from scratch to hold
    # the computer settings and variables.
    if($computerSettings -eq $null) {
        # Create an object to hold the settings
        $computerSettings = ([WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_MachineSettings")).CreateInstance()
        $computerSettings.ResourceID = $computer.ResourceID
        $computerSettings.SourceSite = $siteCode
        $computerSettings.LocaleID = $computer.LocaleID
        $computerSettings.Put()

        # Create an object to hold the variable data
        $computerVariable = ([WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_MachineVariable")).CreateInstance()
        $computerVariable.IsMasked = $isMasked
        $computerVariable.Name = $variableName
        $computerVariable.Value = $variableValue

        # Add the variable to the settings object we created earlier and store it
        $computerSettings.MachineVariables += $computerVariable
        $computerSettings.Put()
    } else {
        # This computer once had variables (or still does), so we need to check for an existing variable with the same name as the one
        # we're trying to add so we can properly overwrite it
        $computerSettings.Get()
        $computerVariables = $computerSettings.MachineVariables

        if($computerVariables -eq $null) {
            # Looks like it held variables in the past but doesn't have them anymore.
            $temporarycomputerSettings = ([WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_MachineSettings")).CreateInstance()
            $computerVariables = $temporarycomputerSettings.MachineVariables

            $computerVariable = ([WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_MachineVariable")).CreateInstance()
            $computerVariable.IsMasked = $isMasked
            $computerVariable.Name = $variableName
            $computerVariable.Value = $variableValue

            $computerVariables += $computerVariable

            $computerSettings.MachineVariables = $computerVariables
            $computerSettings.Put()
        } else {
            # Looks like the computer still has some variables. We'll search through the list of variables to figure out if the computer 
            # already has the variable that we're trying to set.
            $foundExistingVariable = $false
            for($index = 0; ($index -lt $computerVariables.Count) -and ($foundExistingVariable -ne $true); $index++) {
                if($computerVariables[$index].Name -eq $variableName) {
                    # Looks like the computer already has a variable with the same name as the one we're trying to set,
                    # so we're just going to overwrite it.
                    $computerVariables[$index].Value = $variableValue
                    $computerVariables[$index].IsMasked = $isMasked

                    $computerSettings.MachineVariables = $computerVariables
                }
            }

            if($foundExistingVariable -ne $true) {
                # We looked for an existing variable with a matching name, but didn't find it.  So we create a new one from scratch.
                $computerVariable = ([WMIClass]("\\$siteServer\root\sms\site_" + "$siteCode" + ":SMS_MachineVariable")).CreateInstance()
                $computerVariable.IsMasked = $isMasked
                $computerVariable.Name = $variableName
                $computerVariable.Value = $variableValue

                $computerSettings.MachineVariables += $computerVariable
            }
            $computerSettings.Put() 
        }
    }
    return $computerSettings
}

<#
.SYNOPSIS
Returns a list of all computer variables for a specific computer.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and returns a list of all variables associated with that computer.
#>
Function Get-SCCMComputerVariables {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    $computerSettings = Get-WMIObject -computername $siteServer -namespace "root\sms\site_$siteCode" -class "SMS_MachineSettings" | where {$_.ResourceID -eq $computer.ResourceID}
    $computerSettings.Get()

    return $computerSettings.MachineVariables
}

<#
.SYNOPSIS
Deletes a variable for an SCCM computer.

.DESCRIPTION
Takes in information about a specific site, along with a computer name and variable name and attempts to delete that variable.
#>
Function Remove-SCCMComputerVariable {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$siteServer,
        [parameter(Mandatory=$true)][string]$siteCode,
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$variableName
    )

    $computer = Get-SCCMComputer $siteServer $siteCode $computerName
    $computerSettings = Get-WMIObject -computername $siteServer -namespace "root\sms\site_$siteCode" -class "SMS_MachineSettings" | where {$_.ResourceID -eq $computer.ResourceID}
    $computerSettings.Get() 
    
    # Since Powershell arrays are immutable, we create a new one to hold all the computer variable minus the one we're being asked to remove.
    $newVariables = @()
    $computerVariables = $computerSettings.MachineVariables
    for($index = 0; $index -lt $computerVariables.Count; $index++) {
        if($computerVariables[$index].Name -ne $variableName) {
            $newVariables += $computerVariables[$index]
        }
    }
    $computerSettings.MachineVariables = $newVariables
    $computerSettings.Put()
    return $computerSettings.MachineVariables
}

<#
.SYNOPSIS
Allows triggering of SCCM client actions.

.DESCRIPTION
Takes in a computer name and an action id, and attempts to trigger it remotely.

.PARAMETER scheduleId
The available schedule ids are:

HardwareInventory
SoftwareInventory
DataDiscovery
MachinePolicyRetrievalEvalCycle
MachinePolicyEvaluation
RefreshDefaultManagementPoint
RefreshLocation
SoftwareMeteringUsageReporting
SourcelistUpdateCycle
RefreshProxyManagementPoint
CleanupPolicy
ValidateAssignments
CertificateMaintenance
BranchDPScheduledMaintenance
BranchDPProvisioningStatusReporting
SoftwareUpdateDeployment
StateMessageUpload
StateMessageCacheCleanup
SoftwareUpdateScan
SoftwareUpdateDeploymentReEval
OOBSDiscovery
#>
Function Invoke-SCCMClientAction {           
    [CmdletBinding()]             
    param(             
        [parameter(Mandatory=$true)][string]$computerName, 
        [Parameter(Mandatory=$true)][string]$scheduleId    
    )             

    $scheduleIdList = @{
        HardwareInventory = "{00000000-0000-0000-0000-000000000001}";
        SoftwareInventory = "{00000000-0000-0000-0000-000000000002}";
        DataDiscovery = "{00000000-0000-0000-0000-000000000003}";
        MachinePolicyRetrievalEvalCycle = "{00000000-0000-0000-0000-000000000021}";
        MachinePolicyEvaluation = "{00000000-0000-0000-0000-000000000022}";
        RefreshDefaultManagementPoint = "{00000000-0000-0000-0000-000000000023}";
        RefreshLocation = "{00000000-0000-0000-0000-000000000024}";
        SoftwareMeteringUsageReporting = "{00000000-0000-0000-0000-000000000031}";
        SourcelistUpdateCycle = "{00000000-0000-0000-0000-000000000032}";
        RefreshProxyManagementPoint = "{00000000-0000-0000-0000-000000000037}";
        CleanupPolicy = "{00000000-0000-0000-0000-000000000040}";
        ValidateAssignments = "{00000000-0000-0000-0000-000000000042}";
        CertificateMaintenance = "{00000000-0000-0000-0000-000000000051}";
        BranchDPScheduledMaintenance = "{00000000-0000-0000-0000-000000000061}";
        BranchDPProvisioningStatusReporting = "{ 00000000-0000-0000-0000-000000000062}";
        SoftwareUpdateDeployment = "{00000000-0000-0000-0000-000000000108}";
        StateMessageUpload = "{00000000-0000-0000-0000-000000000111}";
        StateMessageCacheCleanup = "{00000000-0000-0000-0000-000000000112}";
        SoftwareUpdateScan = "{00000000-0000-0000-0000-000000000113}";
        SoftwareUpdateDeploymentReEval = "{00000000-0000-0000-0000-000000000114}";
        OOBSDiscovery = "{00000000-0000-0000-0000-000000000120}";
    }
     
    return Invoke-SCCMClientSchedule $computerName $scheduleIdList.$scheduleId
} 

<#
.SYNOPSIS
Allows triggering of SCCM client scheduled messages.

.DESCRIPTION
Takes in a computer name and an scheduled message id, and attempts to trigger it remotely.
#>
Function Invoke-SCCMClientSchedule {           
    [CmdletBinding()]             
    param(             
        [parameter(Mandatory=$true)][string]$computerName, 
        [Parameter(Mandatory=$true)][string]$scheduleId    
    ) 

    $sccmClient = [WMIclass]"\\$computerName\root\ccm:SMS_Client" 
    return $sccmClient.TriggerSchedule($scheduleId)
}

<#
.SYNOPSIS
Contacts a client computer to obtain its advertisement history.

.DESCRIPTION
Takes in a computer name in order to contact a computer and obtain its advertisement history containing active, disabled, and expired advertisements.
#>
Function Get-SCCMClientAdvertisementHistoryForComputer {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$computerName
    )

    return Get-WMIObject -Computer $computerName -Namespace "root\CCM\Policy\Machine\ActualConfig" -Query "Select * from CCM_SoftwareDistribution"
}

<#
.SYNOPSIS
Returns the schedule ID for a specific advertisement on a specific client.

.DESCRIPTION
Takes in a computer name and an advertisement ID and contacts a client to find out the schedule ID for the advertisement.
Useful when trying to trigger and advertisement on demand.
#>
Function Get-SCCMClientAdvertisementScheduleId {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$advertisementId
    )

    $scheduledMessages = Get-WMIObject -Computer $computerName -Namespace "root\CCM\Policy\Machine\ActualConfig" -Query "Select * from CCM_Scheduler_ScheduledMessage"  
    foreach($scheduledMessage in $scheduledMessages) {
        if($($scheduledMessage.ScheduledMessageID).Contains($advertisementId)) {
            return $scheduledMessage.ScheduledMessageID
        }
    }
    return $null
}

<#
.SYNOPSIS
Returns the ssite code of a specific client computer.

.DESCRIPTION
Takes in a computer name and contacts the client to determine its assigned site code.
#>
Function Get-SCCMClientAssignedSite {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$computerName
    )

    $sccmClient = [WMIclass]"\\$computerName\root\ccm:SMS_Client" 
    return $($sccmClient.GetAssignedSite()).sSiteCode
}

<#
.SYNOPSIS
Assignes a new site code to a specific client computer.

.DESCRIPTION
Takes in a computer name and contacts the client to set its assigned site code.
#>
Function Set-SCCMClientAssignedSite {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$computerName,
        [parameter(Mandatory=$true)][string]$siteCode
    )

    $sccmClient = [WMIclass]"\\$computerName\root\ccm:SMS_Client" 
    return $sccmClient.SetAssignedSite($siteCode)
}

<#
.SYNOPSIS
Utility function to convert SCCM date strings into something readable.

.DESCRIPTION
This function exists for convenience so users of the module do not have to try to figure out, if they are not familiar with it, ways
to convert SCCM date strings into something readable.
#>
Function Convert-SCCMDate {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$date
    )

    [System.Management.ManagementDateTimeconverter]::ToDateTime($date);
}

Export-ModuleMember New-SCCMComputer
Export-ModuleMember Remove-SCCMComputer
Export-ModuleMember Get-SCCMComputer
Export-ModuleMember Add-SCCMComputerToCollection
Export-ModuleMember Remove-SCCMComputerFromCollection
Export-ModuleMember Get-SCCMCollection
Export-ModuleMember Get-SCCMCollectionsForComputer
Export-ModuleMember Get-SCCMAdvertisementsForCollection
Export-ModuleMember Get-SCCMAdvertisementsForComputer 
Export-ModuleMember Get-SCCMAdvertisementStatusForComputer
Export-ModuleMember Set-SCCMComputerVariable
Export-ModuleMember Get-SCCMComputerVariables
Export-ModuleMember Remove-SCCMComputerVariable
Export-ModuleMember Invoke-SCCMClientAction
Export-ModuleMember Invoke-SCCMClientSchedule
Export-ModuleMember Get-SCCMClientAdvertisementHistoryForComputer
Export-ModuleMember Get-SCCMClientAdvertisementScheduleId
Export-ModuleMember Get-SCCMClientAssignedSite
Export-ModuleMember Set-SCCMClientAssignedSite
Export-ModuleMember Convert-SCCMDate