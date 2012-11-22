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
        [parameter(Mandatory=$true, Position=0)][string]$computerName, 
        [Parameter(Mandatory=$true, Position=1)][string]$scheduleId    
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
        [parameter(Mandatory=$true, Position=0)][string]$computerName, 
        [Parameter(Mandatory=$true, Position=1)][string]$scheduleId    
    ) 

    $sccmClient = [WMIclass]"\\$computerName\root\ccm:SMS_Client" 
    return $sccmClient.TriggerSchedule($scheduleId)
}

<#
.SYNOPSIS
Contacts a client computer to obtain its software distribution history.

.DESCRIPTION
Takes in a computer name and attempts to contact it to obtain information about its software distribution history.

The information comes from the CCM_SoftwareDistribution WMI class which according to the Microsoft documentation is a "class
that stores information specific to a software distribution, a combination of the properties for the package, program, and advertisement 
that were created to distribute the software."

.PARAMETER computerName
The name of the client to be contacted in order to retrieve the advertisement history.

.EXAMPLE
Get-SCCMSoftwareDistributionHistory -computerName MYCOMPUTER

.NOTES
http://msdn.microsoft.com/en-us/library/cc145304.aspx
#>
Function Get-SCCMClientSoftwareDistributionHistory {
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
        [parameter(Mandatory=$true, Position=0)][string]$computerName,
        [parameter(Mandatory=$true, Position=1)][string]$advertisementId
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
        [parameter(Mandatory=$true, Position=0)][string]$computerName,
        [parameter(Mandatory=$true, Position=1)][string]$siteCode
    )

    $sccmClient = [WMIclass]"\\$computerName\root\ccm:SMS_Client" 
    return $sccmClient.SetAssignedSite($siteCode)
}

<#
.SYNOPSIS
Retrieves a client computer's cache size.

.DESCRIPTION
Contacts a client computer to retrieve information about its cache size.  The value returned represents the cache size in MB.

.PARAMETER computerName
Name of the computer to be contacted.

.EXAMPLE
Get-SCCMClientCacheSize -computerName MYCOMPUTER

Description
-----------
Returns the client computer's cache size in MB.
#>
Function Get-SCCMClientCacheSize {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)][string]$computerName
    )

    $cacheConfig = Get-WMIObject -Computer $computerName -Namespace "root\CCM\SoftMgmtAgent" -Query "Select * From CacheConfig"  
    return $cacheConfig.Size
}

<#
.SYNOPSIS
Changes the cache size for a client computer.

.DESCRIPTION
Contacts a client computer to change its cache size.  The value won't be picked up by the SCCM client on the target computer
until the CcmExec service is restarted.  This function does not attempt to restart the service.

.PARAMETER computerName
Name of the computer to be contacted.

.PARAMETER cacheSize
The new size of the computer's cache in MB.  This value must be greater than 0.

.EXAMPLE
Set-SCCMClientCacheSize -computerName MYCOMPUTER -cacheSize 1000

Description
-----------
Sets the client computer's cache size to 1000MB.
#>
Function Set-SCCMClientCacheSize {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, Position=0)][string]$computerName,
        [parameter(Mandatory=$true, Position=1)][int]$cacheSize
    )

    if($cacheSize -gt 0) {
        $cacheConfig = Get-WMIObject -Computer $computerName -Namespace "root\CCM\SoftMgmtAgent" -Query "Select * From CacheConfig"  
        $cacheConfig.Size = $cacheSize
        $cacheConfig.Put()
    } else {
        Throw "Cache size needs to be greater than 0"
    }
}