<#
.SYNOPSIS
Command line interface for an assortment of SCCM operations.

.DESCRIPTION
The functions in this module provide a command line and scripting interface for automating management of SCCM 2007 environments.

.NOTES
File Name  : SCCM.psm1  
Author     : Andre Bocchini <andrebocchini@gmail.com>  
Requires   : PowerShell 3.0

.LINK
https://github.com/andrebocchini/SCCM-Powershell-Automation-Module
#>

<#
.SYNOPSIS
Attempts to discover the local computer's site provider.

.DESCRIPTION
This function is not exported and is private to this module.  When a user does not specify a site provider for a function that requires that information,
we call this code to try to determine who the provider is automatically.  If we cannot find it, we throw an exception.
#>
Function Get-SCCMSiteProvider {
    $ErrorActionPreference = "Stop"
    try {
        # We need to find this client's management point
        $ccmAuthorityInfo = Get-WMIObject -Namespace "root\ccm" -Class "CCM_Authority"
        foreach($item in $ccmAuthorityInfo) {
            if($item.__CLASS -eq "SMS_Authority") {
                $currentManagementPoint = $item.CurrentManagementPoint
            }
        }     

        # Now we ask the management point for the site provider
        if($currentManagementPoint) {
            
            $providerLocations = Get-WMIObject -ComputerName $currentManagementPoint -Namespace "root\sms" -Query "Select * From SMS_ProviderLocation"
            foreach($location in $providerLocations) {
                if($location.ProviderForLocalSite) { 
                    return $location.Machine
                }
            }
        }
    } catch {
        Throw "Unable to determine site provider.  Please provide one as a parameter."
    } finally{
        $ErrorActionPreference = "Continue"
    }
}

<#
.SYNOPSIS
Attempts to discover the local computer's site code.

.DESCRIPTION
This function is not exported and is private to this module.  When a user does not specify a site code for a function that requires that information,
we call this code to try to determine what the code is automatically.  If we cannot find it, we throw an exception.
#>
Function Get-SCCMSiteCode {
    $ErrorActionPreference = "Stop"
    try {
        # We need to find this client's management point
        $ccmAuthorityInfo = Get-WMIObject -Namespace "root\ccm" -Class "CCM_Authority"

        foreach($item in $ccmAuthorityInfo) {
            if($item.__CLASS -eq "SMS_Authority") {
                $currentManagementPoint = $item.CurrentManagementPoint
            }
        }     
        
        # Now we ask the management point for the site code
        if($currentManagementPoint) {
            $providerLocations = Get-WMIObject -ComputerName $currentManagementPoint -Namespace "root\sms" -Query "Select * From SMS_ProviderLocation" 
            foreach($location in $providerLocations) {
                if($location.ProviderForLocalSite) { 
                    return $location.SiteCode
                }
            }
        }
    } catch {
         Throw "Unable to determine site code.  Please provide one as a parameter."
    } finally{
        $ErrorActionPreference = "Continue"
    }
}

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

    $methodParameters = $site.psbase.GetMethodParameters("ImportMachineEntry")
    $methodParameters.MACAddress = $macAddress
    $methodParameters.NetbiosName = $computerName
    $methodParameters.OverwriteExistingRecord = $false

    $computerCreationResult = $site.psbase.InvokeMethod("ImportMachineEntry", $methodParameters, $null)

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
    return $computer.psbase.Delete()
}

<#
.SYNOPSIS
Returns computers from SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer name or resource ID and will attempt to retrieve an object for the computer.  This function
intentionally ignores obsolete computers by default, but you can specify a parameter to include them.

When invoked without specifying a computer name or resource ID, this function returns a list of all computers found on the specified SCCM site.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site where the computer exists.

.PARAMETER computerName
The name of the computer to be retrieved.

.PARAMETER resourceId
The resource ID of the computer to be retrieved.

.PARAMETER includeObsolete
This switch defaults to false.  If you want your results to include obsolete computers, set this to true when calling this function.

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT -computerName MYCOMPUTER -includeObsolete:$true

Description
-----------
Returns any computer whose name matches MYCOMPUTER found on site SIT.

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT -resourceId 1111

Description
-----------
Returns any computer whose resource ID matches 1111 found on site site SIT.

.EXAMPLE
Get-SCCMComputer -siteProvider MYSITEPROVIDER -siteCode SIT

Description
-----------
Returns all computers (excluding obsolete ones) found on site SIT.
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
        [parameter(Position=0)]
        [parameter(ParameterSetName="name")]
        [ValidateLength(1,15)]
        [string]$computerName,
        [parameter(ParameterSetName="id")]
        [ValidateScript( { $_ -gt 0 } )]
        [int]$resourceId,
        [switch]$includeObsolete=$false
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($computerName) {
        $computerList = Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System" | where { $_.Name -like $computerName }
    } elseif($resourceId) {
        $computerList = Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System" | where { $_.ResourceID -eq $resourceId }
    } else {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_R_System"
    }

    if(!$includeObsolete) {
        $listWithoutObsoleteComputers = @()
        foreach($computer in $computerList) {
            if($computer.Obsolete -ne 1) {
                $listWithoutObsoleteComputers += $computer
            }
        }
        return $listWithoutObsoleteComputers
    } else {
        return $computerList
    }
}

<#
.SYNOPSIS
Adds a computer to a collection in SCCM.

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID, and a collection ID and creates a 
direct membership rule for the computer in the specified collection.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER resourceId
Resource ID of the computer to be added to the collection.

.PARAMETER collectionId
ID of the collection where the computer is to be added.
#>
Function Add-SCCMComputerToCollection {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateScript( { $_ -gt 0 } )]
        [parameter(Mandatory=$true, Position=0)]
        [int]$resourceId,
        [ValidateLength(8,8)]
        [parameter(Mandatory=$true, Position=1)]
        [string]$collectionId
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

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collection) {
        Throw "Unable to retrieve collection with ID $collectionId"
    }

    # We want to get a list of all the collection this computer already belongs to so we can check later if it already is
    # a member of the collection passed as a parameter to this function
    $currentCollectionMembershipList = Get-SCCMCollectionsForComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    $currentCollectionIdList = @()
    foreach($currentCollectionMembership in $currentCollectionMembershipList) {
        $currentCollectionIdList += ($currentCollectionMembership.CollectionID).ToUpper()
    }

    # We will only add the computer if it isn't already part of the collection passed as a parameter to this function
    if($currentCollectionIdList -notcontains ($collection.CollectionID).ToUpper()) {
        $collectionRule = New-SCCMCollectionRuleDirect -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId

        $addToCollectionParameters = $collection.GetmethodParameters("AddMembershipRule")
        $addToCollectionParameters.CollectionRule = $collectionRule
        $status = $collection.InvokeMethod("AddMembershipRule", $addToCollectionParameters, $null)
        
        if($status.ReturnValue -ne 0) {
            Throw "Failed to add computer $($computer.Name) to collection $($collection.Name)"
        }
    }
}

<#
.SYNOPSIS
Removes a computer from a specific collection

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID, and a collection ID and removes the computer from the collection.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER resourceId
Resource ID of the computer to be removed from the collection.

.PARAMETER collectionId
ID of the collection where the computer is a member.
#>
Function Remove-SCCMComputerFromCollection {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,

        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [int]$resourceId,
        [ValidateLength(8,8)]
        [parameter(Mandatory=$true, Position=1)]
        [string]$collectionId
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

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collection) {
        Throw "Unable to retrieve collection with ID $collectionId"
    }

    $collectionRule = New-SCCMCollectionRuleDirect -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    $collectionRule.ResourceID = $resourceId
    
    $status = $collection.DeleteMembershipRule($collectionRule)
    if($status.ReturnValue -ne 0) {
        Throw "Failed to remove computer $($computer.Name) from collection $($collection.Name)"
    }
}

<#
.SYNOPSIS
Creates a direct collection rule.

.DESCRIPTION
Creates a direct collection rule.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER resourceId
Resource ID of the computer that is part of the rule.

.LINK
http://msdn.microsoft.com/en-us/library/cc145537.aspx
#>
Function New-SCCMCollectionRuleDirect {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateScript( { $_ -gt 0 } )]
        [parameter(Mandatory=$true, Position=0)]
        [int]$resourceId
    ) 
    
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $rule = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_CollectionRuleDirect")).CreateInstance()
    if($rule) {
        $rule.ResourceClassName = "SMS_R_System"
        $rule.ResourceID = $resourceId       
    }
    return $rule
}

<#
.SYNOPSIS
Creates static SCCM collections

.DESCRIPTION
Allows the creation of static collections.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code where the collection is to be created.

.PARAMETER collectionName
The name of the new collection.

.PARAMETER parentCollectionId
If the static collection is not bound to a parent, it will not show up in the console.  This parameter is mandatory and must be valid.

.PARAMETER collectionComment
Optional parameter.  Attaches a comment to the collection.

.EXAMPLE
New-SCCMStaticCollection -siteProvider MYSITEPROVIDER -siteCode SIT -collectionName MYCOLLECTIONNAME -parentCollectionId SIT00012 -collectionComment "This is a comment"
#>
Function New-SCCMStaticCollection {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$collectionName,
        [parameter(Mandatory=$true, Position=1)][string]$parentCollectionId,
        [string]$collectionComment
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if(Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $parentCollectionId) {
        if(!(Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionName $collectionName)) {
            $newCollection = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_Collection")).CreateInstance()
            $newCollection.Name = $collectionName
            $newCollection.Comment = $collectionComment
            $newCollection.OwnedByThisSite = $true

            $newCollection.Put() | Out-Null

            # Now we establish the parent to child relationship of the two collections. If we create a collection without
            # establishing the relationship, the new collection will not be visible in the console.
            $newCollectionId = (Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionName $collectionName).CollectionID
            $newCollectionRelationship = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_CollectToSubCollect")).CreateInstance()
            $newCollectionRelationship.parentCollectionID = $parentCollectionId
            $newCollectionRelationship.subCollectionID = $newCollectionId

            $newCollectionRelationship.Put() | Out-Null

            return Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $newCollectionId
        } else {
            Throw "A collection named $collectionName already exists"
        }
    } else {
        Throw "Invalid parent collection"
    }
}

<#
.SYNOPSIS
Returns a collection refresh schedule.

.DESCRIPTION
Returns a collection refresh schedule.  If the collection is set to refresh manually, it returns null.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER collectionId
ID of the collection whose collection is being returned.

.LINK
http://msdn.microsoft.com/en-us/library/cc145320.aspx
#>
Function Get-SCCMCollectionRefreshSchedule {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateLength(8,8)][string]$collectionId
    )

    $refreshTypeManual = 1
    $refreshTypeAuto = 2

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if($collection) {
        $collection.Get() | Out-Null
        if($collection.RefreshType -eq $refreshTypeAuto) {
            return $collection.RefreshSchedule
        }         
    }
}

<#
.SYNOPSIS
Sets a collection refresh schedule.

.DESCRIPTION
Sets a collection refresh schedule. 

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER collectionId
ID of the collection whose collection whose schedule is being set.

.PARAMETER refreshType
Values allowed are 1 for manual refresh, and 2 for scheduled refresh.  If 2 is specified, a schedule must
be passed as a parameter.

.PARAMETER refreshSchedule
A schedule token representing the collection refresh schedule.

.LINK
http://msdn.microsoft.com/en-us/library/cc145320.aspx
#>
Function Set-SCCMCollectionRefreshSchedule {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateLength(8,8)][string]$collectionId,
        [parameter(Mandatory=$true, Position=1)][ValidateRange(1,2)][int]$refreshType = 1,
        [parameter(Position=2)][ValidateScript( { !(!$_ -and $refreshType -eq 2) } )]$refreshSchedule
    )
    
    if($refreshType -eq 2) {
        if(!($PSBoundParameters.refreshSchedule)) {
            Throw "No refresh schedule specified"
        }
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collection) {
        Throw "Unable to retrieve collection with ID $collectionId"
    } else {
        $collection.Get() | Out-Null
    }

    $collection.RefreshType = $refreshType
    if($refreshType -eq 2) {
        $collection.RefreshSchedule = $refreshSchedule
    }
    $collection.Put() | Out-Null
}

<#
.SYNOPSIS
Saves a collection back into the SCCM database.

.DESCRIPTION
This function is used to save direct property changes to collections back to the SCCM database.

.PARAMETER collection
The collection object to be put back into the database.
#>
Function Save-SCCMCollection {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$collection
    )

    $collection.Put() | Out-Null
}

<#
.SYNOPSIS
Deletes SCCM collections

.DESCRIPTION
Takes in information about a specific site, along with a collection ID and deletes it.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site to be queried.

.PARAMETER collectionId
The id of the collection to be deleted.

.EXAMPLE
Remove-SCCMCollection -siteProvider MYSITEPROVIDER -siteCode SIT -collectionId MYID

Description
-----------
Deletes the collection with id MYID from site SIT on MYSITEPROVIDER.
#>
Function Remove-SCCMCollection {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=1)][string]$collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    return $collection.psbase.Delete()
}

<#
.SYNOPSIS
Retrieves SCCM collection objects from the specified site.

.DESCRIPTION
Takes in information about a specific site and a collection name and returns an object representing that collection.  If no collection name is specified, it returns all collections found on the specified site.

.PARAMETER siteProvider
The name of the site.

.PARAMETER siteCode
The 3-character site code for the site where the collection exists.

.PARAMETER collectionName
Optional parameter.  If specified, the function will try to find a collection that matches the name provided.

.PARAMETER collectionId
Optional parameter.  If specified, the function will try to find a collection that matches the id provided.

.EXAMPLE
Get-SCCMCollection -siteProvider MYSITEPROVIDER -siteCode SIT -collectionName MYCOLLECTION

Description
-----------
Retrieve the collection named MYCOLLECTION from site SIT

.EXAMPLE
Get-SCCMCollection -siteProvider MYSITEPROVIDER -siteCode SIT

Description
-----------
Retrieve all collections from site SIT

.EXAMPLE
Get-SCCMCollection -siteProvider MYSITEPROVIDER -siteCode SIT | Select-Object Name,CollectionID

Description
-----------
Retrieve all collections from site SIT and filter out only their names and IDs
#>
Function Get-SCCMCollection {
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
        [parameter(ParameterSetName="name", Position=0)]
        [string]$collectionName,
        [parameter(ParameterSetName="id", Position=1)]
        [string]$collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($collectionName) {
        return Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Collection" | Where { $_.Name -like $collectionName } 
    } elseif($collectionId) {
        return Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Collection" | Where { $_.CollectionID -eq $collectionId }
    } else {
        return Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Collection"
    }
}

<#
.SYNOPSIS
Retrieves the list of members of a collection.

.DESCRIPTION
Takes in information about an SCCM site along with a collection ID and returns a list of all members of the target collection.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
Site code for the site containing the collection.

.PARAMETER collectionId
The ID of the collection whose members are being retrieved.

.EXAMPLE
Get-SCCMCollectionMembers -siteProvider MYSITEPROVIDER -siteCode SIT -collectionId SIT00012

.EXAMPLE
Get-SCCMCollectionMembers -siteProvider MYSITEPROVIDER -siteCode SIT -collectionId SIT00012 | Select-Object -ExpandProperty Name

Description
-----------
Retrieves all members of collection SIT00012 and lists only their names
#>
Function Get-SCCMCollectionMembers {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_CollectionMember_a" | Where { $_.CollectionID -eq $collectionId }
}

<#
.SYNOPSIS
Returns a list of collection IDs that a computer belongs to.

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID and returns an array containing the collection IDs that the computer belongs to.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character code for the site where the computer exists.

.PARAMETER resourceId
Resource ID for the computer whose list of collections are being retrieved.
#>
Function Get-SCCMCollectionsForComputer {
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

    # First we find all membership associations that match the colection ID of the computer in question
    $computer = Get-SCCMComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId

    if(!$computer) {
        Throw "Unable to retrieve computer with resource ID $resourceId"
    }

    $computerCollectionIds = @()
    $collectionMembers = Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_CollectionMember_a Where ResourceID= $($computer.ResourceID)"
    foreach($collectionMember in $collectionMembers) {
        $computerCollectionIds += $collectionMember.CollectionID
    }

    # Now that we have a list of collection IDs, we want to retrieve and return some rich collection objects
    $allServerCollections = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode
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
Creates a new SCCM advertisement.

.DESCRIPTION
Creates a new SCCM advertisement for a software distribution program and assigns it to a collection.  

This function provides a limited set of parameters so it can create a basic package.  It will create the object, save it to the 
database and return a copy of it so you can finish customizing it.  Once you finish customizing the properties of the object, save 
it back using Save-SCCMAdvertisement.  Follow the link in the LINK section of this documentation block to find out the options available 
for customizing an advertisement.

.PARAMETER siteProvider
Site provider for the site where the advertisement will be created.

.PARAMETER siteCode
The 3-character code for the site where the advertisement will be created.

.PARAMETER advertisementName
Name of the new advertisement.

.PARAMETER collectionId
Collection ID of the collection where the advertisement will be assigned.

.PARAMETER packageId
ID of the package to be advertised.

.PARAMETER programName
Named of the program to be advertised that is part of the package definied by the parameter packageId.

.LINK
http://msdn.microsoft.com/en-us/library/cc146108.aspx
#>
Function New-SCCMAdvertisement {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true)][string]$advertisementName,
        [parameter(Mandatory=$true)][string]$collectionId,
        [parameter(Mandatory=$true)][string]$packageId,
        [parameter(Mandatory=$true)][string]$programName
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if(!(Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId)) {
        Throw "Invalid collection with ID $collectionId"
    } elseif(!(Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId)) {
        Throw "Invalid package with ID $packageId"
    } elseif(!(Get-SCCMProgram -siteProvider $siteProvider -siteCode $siteCode $packageId $programName)) {
        Throw "Invalid program with name `"$programName`""
    } else {
        $newAdvertisement = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_Advertisement")).CreateInstance()
        $newAdvertisement.AdvertisementName = $advertisementName
        # For some reason, if you try to pass the collection ID with lowercase characters, the SCCM console will crash when you right-click it and
        # visit the Advertisements tab.  So just to be safe, we convert it to upper case.
        $newAdvertisement.CollectionID = $collectionId.ToUpper()
        $newAdvertisement.Comment = ""
        $newAdvertisement.PackageID = $packageId
        $newAdvertisement.ProgramName = $programName
        $newAdvertisement.PresentTime = Convert-DateToSCCMDate $(Get-Date)
        $newAdvertisement.AdvertFlags = 33554464
        $newAdvertisement.RemoteClientFlags = 8240
        $newAdvertisement.Priority = 2
        $advertisementCreationResult = $newAdvertisement.Put()

        $newAdvertisementId = $($advertisementCreationResult.RelativePath).TrimStart('SMS_Advertisement.AdvertisementID=')
        $newAdvertisementId = $newAdvertisementId.Substring(1,8)

        return Get-SCCMAdvertisement -siteProvider $siteProvider -siteCode $siteCode -advertisementId $newAdvertisementId
    }
}

<#
.SYNOPSIS
Saves an advertisement back into the SCCM database.

.DESCRIPTION
The functions in this module that are used to create advertisements only have a limited number of supported parameters, but 
SCCM advertisements are objects with a large number of settings.  When you create an advertisement, it is likely that you will
want to edit some of those settings.  Once you are finished editing the properties of the advertisement, you need to save it back
into the SCCM database by using this method.

.PARAMETER advertisement
The advertisement object to be put back into the database.
#>
Function Save-SCCMAdvertisement {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$advertisement
    )

    $advertisement.Put() | Out-Null
}

<#
.SYNOPSIS
Deletes an SCCM advertisement.

.DESCRIPTION
Deletes an SCCM advertisement based on an advertisement ID.

.PARAMETER siteProvider
The site provider for the site where the advertisement exists.

.PARAMETER siteCode
The 3-character code for the site where the advertisement exists.

.PARAMETER advertisementId
The ID of the advertisement to be deleted.
#>
Function Remove-SCCMAdvertisement {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$advertisementId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $advertisement = Get-SCCMAdvertisement -siteProvider $siteProvider -siteCode $siteCode -advertisementId $advertisementId
    if($advertisement) {
        $advertisement.psbase.Delete()
    } else {
        Throw "Invalid advertisement with ID $advertisementId"
    }
}

<#
.SYNOPSIS
Retrieves SCCM advertisements from the site provider.

.DESCRIPTION
Takes in information about a specific site and an advertisement name and/or and advertisement ID and returns advertisements matching the specified parameters.  
If no advertisement name or ID is specified, it returns all advertisements found on the site.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site where the advertisement exists.

.PARAMETER advertisementName
Optional parameter.  If specified, the function will try to find advertisements that match the specified name.

.PARAMETER advertisementId
Optional parameter.  If specified, the function will try to find advertisements that match the specified ID.

.EXAMPLE
Get-SCCMAdvertisement -siteProvider MYSITEPROVIDER -siteCode SIT -advertisementName MYADVERTISEMENT

Description
-----------
Retrieve the advertisement named MYADVERTISEMENT from site SIT on MYSITEPROVIDER

.EXAMPLE
Get-SCCMAdvertisement -siteProvider MYSITEPROVIDER -siteCode SIT

Description
-----------
Retrieve all advertisements from site SIT on MYSITEPROVIDER

.EXAMPLE
Get-SCCMAdvertisement -siteProvider MYSITEPROVIDER -siteCode SIT | Select-Object Name,AdvertisementID

Description
-----------
Retrieve all advertisements from site SIT on MYSITEPROVIDER and filter out only their names and IDs
#>
Function Get-SCCMAdvertisement {
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
        [parameter(Position=0)]
        [parameter(ParameterSetName="name")]
        [ValidateNotNull()]
        [string]$advertisementName,
        [parameter(Position=1)]
        [parameter(ParameterSetName="id")]
        [ValidateNotNull()]
        [string]$advertisementId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($advertisementName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Advertisement" | where { $_.AdvertisementName -like $advertisementName }
    } elseif($advertisementId) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Advertisement" | where { $_.AdvertisementID -eq $advertisementId }
    } else { 
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_Advertisement"
    }
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
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WMIObject -Computer $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Advertisement WHERE CollectionID='$collectionId'"
}

<#
.SYNOPSIS
Returns a list of advertisements assigned to a computer.

.DESCRIPTION
Takes in information about a specific site, along with a computer resource ID and returns all advertisements assigned to that computer.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character code for the site where the computer exists.

.PARAMETER resourceId
Resource ID of the computers whose advertisements are being retrieved.
#>
Function Get-SCCMAdvertisementsForComputer {
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

    $collections = Get-SCCMCollectionsForComputer -siteProvider $siteProvider -siteCode $siteCode -resourceId $resourceId
    $computerAdvertisements = @()
    foreach($collection in $collections) {
        $advertisements = Get-SCCMAdvertisementsForCollection -siteProvider $siteProvider -siteCode $siteCode $collection.CollectionID
        foreach($advertisement in $advertisements) {
            $computerAdvertisements += $advertisement
        }
    }
    
    return $computerAdvertisements
}

<#
.SYNOPSIS
Returns the assigned schedule for an advertisement.

.DESCRIPTION
Returns an array containing all of the assigned schedule objects for an advertisement.

.PARAMETER advertisement
The advertisement object whose schedule is being retrieved.

.LINK
http://msdn.microsoft.com/en-us/library/cc145924.aspx
#>
Function Get-SCCMAdvertisementAssignedSchedule {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $advertisement
    )    

    $advertisement.Get() | Out-Null
    return $advertisement.AssignedSchedule
}

<#
.SYNOPSIS
Sets the assigned schedule for an advertisement.

.DESCRIPTION
Sets the assigned schedule for an advertisement using an array of SMS_ScheduleToken objects passed as a parameter.

.PARAMETER advertisement
The advertisement whose schedule is being set.

.PARAMETER schedule
The array of SMS_ScheduleToken objects to be set as the advertisement schedule.  If a a $null or empty are passed, the 
assignment schedule is cleared.

.NOTES
You can create schedule token objects using the schedule token functions in this module.

.LINK
http://msdn.microsoft.com/en-us/library/cc145924.aspx
#>
Function Set-SCCMAdvertisementAssignedSchedule {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        $advertisement,
        [parameter(Mandatory=$true, Position=1)]
        [AllowNull()]
        $schedule
    )

    $advertisement.Get() | Out-Null
    $advertisement.AssignedSchedule = $schedule
    $advertisement.Put() | Out-Null
}

<#
.SYNOPSIS
Returns a list of advertisements for a specific package.

.DESCRIPTION
Takes in information about a specific site, along with a package ID and returns all advertisements for that package.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site to be queried.

.PARAMETER packageId
The ID of the package whose advertisements the function should retrieve.

.EXAMPLE
Get-SCCMAdvertisementForPackage -siteProvider MYSITEPROVIDER -siteCode SIT -packageId MYPACKAGEID

Description
-----------
Retrieve all advertisements from site SIT on MYSITEPROVIDER for package with ID equal to MYPACKAGEID
#>
Function Get-SCCMAdvertisementsForPackage {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$packageId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Advertisement" | where { $_.PackageID -eq $packageId }
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
        [parameter(Mandatory=$true)]$resourceId
    )

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

    $computer = Get-SCCMComputer -siteProvider $siteProvider -siteCode siteCode -resourceId $resourceId
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

<#
.SYNOPSIS
Creates SCCM software distribution packages.

.DESCRIPTION
Takes in information about a specific site along with package details and creates a new software distribution package.  IF
successful, the function return a WMI object for the new package.

This function provides a limited set of parameters so it can create a basic package.  It will create the object, save it to the 
database and return a copy of it so you can finish customizing it.  Once you finish customizing the properties of the object, save 
it back using Save-SCCMPackage.  Follow the link in the LINK section of this documentation block to find out the options available 
for customizing a package.

.PARAMETER siteProvider
The name of the site provider where the package will be created.

.PARAMETER siteCode
The 3-character site code for the site where the package will be created.

.PARAMETER packageName
Name of the new package.

.PARAMETER packageDescription
A description for the new package.

.PARAMETER packageVersion
The new package's version.

.PARAMETER packageManufacturer
The name of the developer of the package being created.

.PARAMETER packageLanguage
The language of the softwre being packaged.

.PARAMETER packageSource
Optional parameter.  If not specified, the package will be created without source files.  This parameter can take
the value of a local path on the site server, or a UNC path for a network share.

.LINK
http://msdn.microsoft.com/en-us/library/cc144959.aspx
#>
Function New-SCCMPackage {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true)][string]$packageName,
        [parameter(Mandatory=$true)][string]$packageDescription,
        [parameter(Mandatory=$true)][string]$packageVersion,
        [parameter(Mandatory=$true)][string]$packageManufacturer,
        [parameter(Mandatory=$true)][string]$packageLanguage,
        [string]$packageSource   
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $newPackage = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_Package")).CreateInstance()
    $newPackage.Name = $packageName
    $newPackage.Description = $packageDescription
    $newPackage.Version = $packageVersion
    $newPackage.Manufacturer = $packageManufacturer
    $newPackage.Language = $packageLanguage
    if($packageSource) {
        $newPackage.PkgSourceFlag = 2
        $newPackage.PkgSourcePath = $packageSource
    } else {
        $newPackage.PkgSourceFlag = 1
    }
    $packageCreationResult = $newPackage.Put()

    if($packageCreationResult) {
        $newPackageIdTokens = $($packageCreationResult.RelativePath).Split("=")
        $newPackageId = $($newPackageIdTokens[1]).TrimStart("`"")
        $newPackageId = $($newPackageId).TrimEnd("`"")

        return Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $newPackageId
    } else {
        Throw "Package creation failed"
    }
}

<#
.SYNOPSIS
Saves a package back into the SCCM database.

.DESCRIPTION
The functions in this module that are used to create packages only have a limited number of supported parameters, but 
SCCM packages are objects with a large number of settings.  When you create a package, it is likely that you will
want to edit some of those settings.  Once you are finished editing the properties of the package, you need to save it back
into the SCCM database by using this method.

.PARAMETER package
The package object to be put back into the database.
#>
Function Save-SCCMPackage {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$package
    )

    $package.Put() | Out-Null
}

<#
.SYNOPSIS
Removes SCCM packages from the specified site.

.DESCRIPTION
Takes in information about a specific site and a package ID and deletes the package.

.PARAMETER siteProvider
The name of the site provider for the site where the package exists.

.PARAMETER siteCode
The 3-character site code for the site where the package exists.

.PARAMETER packageId
The ID of the package to be deleted.

.EXAMPLE
Remove-SCCMPackage -siteProvider MYSITEPROVIDER -siteCode SIT -packageId MYID

Description
-----------
Deletes the package with ID MYID from SIT on MYSITEPROVIDER
#>
Function Remove-SCCMPackage {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$packageId
    )    

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $package = Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId
    if($package) {
        $package.psbase.Delete() | Out-Null
    } else {
        Throw "Invalid package with ID $packageId"
    }
}

<#
.SYNOPSIS
Retrieves SCCM packages from the specified site.

.DESCRIPTION
Takes in information about a specific site and a package ID and returns a package with a matching ID.  If no package ID is specified, it returns all packages found on the specified site.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER packageName
The name of the package being searched for.

.PARAMETER packageId
Optional parameter.  If specified, the function attempts to match the package by package ID.

.EXAMPLE
Get-SCCMPackage -siteProvider MYSITEPROVIDER -siteCode SIT

Description
-----------
Retrieve all packages from site SIT on MYSITEPROVIDER

.EXAMPLE
Get-SCCMPackage -siteProvider MYSITEPROVIDER -siteCode SIT | Select-Object Name,PackageID

Description
-----------
Retrieve all packages from site SIT on MYSITEPROVIDER and filter out only their names and IDs
#>
Function Get-SCCMPackage {
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
        [parameter(Position=0)]
        [parameter(ParameterSetName="name")]
        [ValidateNotNull()]
        [string]$packageName,
        [parameter(ParameterSetName="id")]
        [parameter(Position=1)][string]$packageId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($packageName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Package" | where { ($_.Name -like $packageName) }
    } elseif($packageId) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Package" | where { ($_.PackageID -eq $packageId) }
    } else { 
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_Package"
    }
}

<#
.SYNOPSIS
Creates a new SCCM program.

.DESCRIPTION
Creates a new SCCM program and associates it with a software distribution package.  This function currently only allows
for little customization and creates the program with most default settings.

This function provides a limited set of parameters so it can create a basic program.  It will create the object, save it to the 
database and return a copy of it so you can finish customizing it.  Once you finish customizing the properties of the object, save 
it back using Save-SCCMProgram.  Follow the link in the LINK section of this documentation block to find out the options available 
for customizing a program.

.PARAMETER siteProvider
Site provider for the site where the package containing the new program resides.

.PARAMETER siteCode
Site code for the site where the package containing the new program resides.

.PARAMETER packageId
ID of the package that will contain the new program.

.PARAMETER programName
A unique name for the program.  If this name matches the name for an existing program in the same package,
the program will be overwritten.

.PARAMETER programCommandLine
The command line that will be executed when this program runs.

.LINK
http://msdn.microsoft.com/en-us/library/cc144361.aspx
#>
Function New-SCCMProgram {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true)][string]$packageId,
        [parameter(Mandatory=$true)][string]$programName,
        [parameter(Mandatory=$true)][string]$programCommandLine
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if(Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId) { 
        $newProgram = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_Program")).CreateInstance()
        $newProgram.ProgramName = $programName
        $newProgram.PackageID = $packageId
        $newProgram.CommandLine = $programCommandLine

        $programCreationResult = $newProgram.Put()
        if($programCreationResult) {
            $newProgramId = $($programCreationResult.RelativePath).TrimStart('SMS_Program.PackageID=')
            $newProgramId = $newProgramId.Substring(1,8)

            return Get-SCCMProgram -siteProvider $siteProvider -siteCode $siteCode $packageId $programName
        } else {
            Throw "Program creation failed"
        }
    } else {
        Throw "Invalid package ID $packageId"
    }
}

<#
.SYNOPSIS
Saves a program back into the SCCM database.

.DESCRIPTION
The functions in this module that are used to create programs only have a limited number of supported parameters, but 
SCCM programs are objects with a large number of settings.  When you create a program, it is likely that you will
want to edit some of those settings.  Once you are finished editing the properties of the program, you need to save it back
into the SCCM database by using this method.

.PARAMETER program
The program object to be put back into the database.
#>
Function Save-SCCMProgram {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$program
    )

    $program.Put() | Out-Null
}

<#
.SYNOPSIS
Deletes a SCCM program.

.DESCRIPTION
Deletes a specific program belonging to specific software distribution package.

.PARAMETER siteProvider
Site provider for the site where the package containing the program resides.

.PARAMETER siteCode
Site code for the site where the package containing the program resides.

.PARAMETER packageId
ID of the package that will contains the program to be removed.

.PARAMETER programName
Name of the program to be deleted.
#>
Function Remove-SCCMProgram {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$packageId,
        [parameter(Mandatory=$true, Position=1)][string]$programName
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $program = Get-SCCMProgram -siteProvider $siteProvider -siteCode $siteCode $packageId $programName
    if($program) {
        return $program.psbase.Delete()
    } else {
        Throw "Invalid program ID or program name"
    }
}

<#
.SYNOPSIS
Retrieves SCCM programs from the specified site.

.DESCRIPTION
Takes in information about a specific site and a package ID and returns a program matching the package ID and program name
passed as parameters.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code for the site.

.PARAMETER packageId
Optional parameter.  If specified along with a program name, the function returns a program matching the and the package ID specified.
If only a package ID is specified, the function returns all programs that belong to the package.  If absent, the function returns all programs for the site.

.PARAMETER programName
Name of the program to be returned.  If specified, a package ID must also be specified.

.EXAMPLE
Get-SCCMProgram -siteProvider MYSITEPROVIDER -siteCode SIT -packageId PACKAGEID

Description
-----------
Retrieve all programs with package ID PACKAGEID from site SIT on MYSITEPROVIDER

.EXAMPLE
Get-SCCMProgram -siteProvider MYSITEPROVIDER -siteCode SIT

Description
-----------
Retrieve all programs from site SIT on MYSITEPROVIDER

.EXAMPLE
Get-SCCMProgram -siteProvider MYSITEPROVIDER -siteCode SIT | Select-Object ProgramName,PackageID

Description
-----------
Retrieve all programs from site SIT on MYSITEPROVIDER and filter out only their names and package IDs
#>
Function Get-SCCMProgram {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Position=0)][string]$packageId,
        [parameter(Position=1)][string]$programName
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($packageId -and $programName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Program" | where { ($_.PackageID -eq $packageId) -and ($_.ProgramName -eq $programName) }
    } elseif($packageId -and !$programName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Program" | where { ($_.PackageID -eq $packageId) }
    } elseif(!$packageId -and $programName) {
        Throw "If a program name is specified, a package ID must also be specified"
    } else { 
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_Program"
    }
}

<#
.SYNOPSIS
Adds a package to one or more distribution points.

.DESCRIPTION
Adds a package to one or more distribution points.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
The 3-character code for the site containing the disribution points.

.PARAMETER packageId
The ID of the package that needs to be added to a distribution point.

.PARAMETER distributionPointList
A list of one or more distribution points that will have the package added to them.  These are WMI
objects with information about distribution points.  The easiest way to obtain them is to use Get-SCCMDistributionPoints.

.EXAMPLE
$dpList = Get-SCCMDistributionPoint -siteProvider MYSITEPROVIDER -siteCode SIT
Add-SCCMPackageToDistributionPoint -siteProvider MYSITEPROVIDER -siteCode SIT -packageId SIT00000 -distributionPointList $dpList

Description
-----------
This will add the package SIT00000 to every distribution point on the site SIT.
#>
Function Add-SCCMPackageToDistributionPoint {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$packageId,
        [parameter(Mandatory=$true, Position=1)]$distributionPointList
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if(Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId) {
        foreach($distributionPoint in $distributionPointList) {
            $newDistributionPoint = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_DistributionPoint")).CreateInstance()
            $newDistributionPoint.ServerNALPath = $distributionPoint.NALPath
            $newDistributionPoint.PackageID = $packageId
            $newDistributionPoint.SiteCode = $distributionPoint.SiteCode
            $newDistributionPoint.SiteName = $distributionPoint.SiteCode
            $newDistributionPoint.psbase.Put() | Out-Null
        }
    } else {
        Throw "Invalid package ID $packageId"
    }
}

<#
.SYNOPSIS
Removes a package from one or more distribution points.

.DESCRIPTION
Removes a package from one or more distribution points.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
The 3-character code for the site containing the disribution points.

.PARAMETER packageId
The ID of the package to be removed from distribution points.

.PARAMETER distributionPointList
A list of one or more distribution points that will have the package removed from them.  These are WMI
objects with information about distribution points.  The easiest way to obtain them is to use Get-SCCMDistributionPoints.

.EXAMPLE
$dpList = Get-SCCMDistributionPoint -siteProvider MYSITEPROVIDER -siteCode SIT
Remove-SCCMPackageFromDistributionPoint -siteProvider MYSITEPROVIDER -siteCode SIT -packageId SIT00000 -distributionPointList $dpList

Description
-----------
This will remove the package SIT00000 from every distribution point on the site SIT.
#>
Function Remove-SCCMPackageFromDistributionPoint {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][string]$packageId,
        [parameter(Mandatory=$true, Position=1)]$distributionPointList
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if(Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId) {
        foreach($distributionPoint in $distributionPointList) {
            $distributionPointToBeDeleted = Get-WmiObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_DistributionPoint" | Where { ($_.ServerNALPath -eq $distributionPoint.NALPath) -and ($_.PackageID -eq $packageId) }
            if($distributionPointToBeDeleted) {
                $distributionPointToBeDeleted.psbase.Delete()
            }
        }
    } else {
        Throw "Invalid package ID $packageId"
    }
}

<#
.SYNOPSIS
Returns a list of distribution points for a particular site.

.DESCRIPTION
Takes in information about a particular site and returns WMI objects for each distribution point it finds.

.PARAMETER siteProvider
The site provider.

.PARAMETER siteCode
The 3-character code for the site that holds the distribution points.

.EXAMPLE
Get-SCCMDistributionPoints -siteProvider MYSITEPROVIDER -siteCode SIT
#>
Function Get-SCCMDistributionPoints {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode
    )    

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WmiObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_SystemResourceList" | Where { ($_.RoleName -eq "SMS Distribution Point") -and  ($_.SiteCode -eq $siteCode) }
}

<#
.SYNOPSIS
Returns the maintenance windows for a collection.

.DESCRIPTION
Returns the maintenance windows for a collection.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER collectionId
ID of the collection whose maintenance windows are being retrieved.

.LINK
http://msdn.microsoft.com/en-us/library/cc143300.aspx
#>
Function Get-SCCMMaintenanceWindows {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateNotNull()][string]$collectionId
    )    

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collectionSettings = GET-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_CollectionSettings" | Where { $_.CollectionID -eq $collectionId }
    if($collectionSettings) {
        $collectionSettings.Get()
        return $collectionSettings.ServiceWindows
    }
}

<#
.SYNOPSIS
Retrieves the schedules from a maintenance window object.

.DESCRIPTION
Retrieves the schedules from a maintenance window object.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER maintenanceWindow
Maintenance window object whose schedules are being retrieved.

.LINK
http://msdn.microsoft.com/en-us/library/cc145924.aspx
#>
Function Get-SCCMMaintenanceWindowSchedules {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateNotNull()]$maintenanceWindow
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($maintenanceWindow) {
        $schedules = $maintenanceWindow.ServiceWindowSchedules

        if($schedules) {
            $scheduleMethod = [WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ScheduleMethods")
            $result = $scheduleMethod.ReadFromString($schedules)
            return $result.TokenData
        }
    }
}

<#
.SYNOPSIS
Gets a list of supported platforms for programs for the specified site.

.DESCRIPTION
Gets a list of supported platforms for programs for the specified site.  This is useful when creaating programs and needing
to determine the supported platforms one can configure for those programs.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.EXAMPLE
Get-SCCMSupported Platforms -siteProvider MYSITEPROVIDER -siteCode SIT

Description
-----------
Gets a list of all supported platforms on the site.

.LINK
http://msdn.microsoft.com/en-us/library/cc144734.aspx
#>
Function Get-SCCMSupportedPlatforms {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WmiObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_SupportedPlatforms"
}

<#
.SYNOPSIS
Creates a new supported platform object.

.DESCRIPTION
Creates a new supported platform object used to configure supported platforms for SCCM programs.  The values passed
to this function have to match values obtained from Get-SCCMSupportedPlatforms.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER name
The display name of the operating system.

.PARAMETER maxVersion
The maximum version of the operating system.

.PARAMETER minVersion
The minimum version of the operating system.

.PARAMETER platform
The platform the operating system runs on (x86, x64, IA64).

.EXAMPLE
New-SCCMSupportedPlatform -siteProvider MYSITEPROVIDER -siteCode SIT -name "WinNT" -maxVersion "6.10.999.0" -minVersion "6.10.7600.0" -platform "i386"

.LINK
http://msdn.microsoft.com/en-us/library/cc146485.aspx
#>
Function New-SCCMSupportedPlatform {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true)][string]$name,
        [parameter(Mandatory=$true)][string]$maxVersion,
        [parameter(Mandatory=$true)][string]$minVersion,
        [parameter(Mandatory=$true)][string]$platform
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $newPlatform = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_OS_Details")).CreateInstance()
    if($newPlatform) {
        $newPlatform.Name = $name
        $newPlatform.MaxVersion = $maxVersion
        $newPlatform.MinVersion = $minVersion
        $newPlatform.Platform = $platform

        return $newPlatform
    } else {
        Throw "Unable to create new supported platform"
    }
}

<#
.SYNOPSIS
Retrieves the list of supported platforms for a program.

.DESCRIPTION
The list of supported operating systems for a program is a lazy property and does not get retrieved when a program is
retrieved via a WMI query; it has to be explicitly requested.

.PARAMETER program
The program whose list of supported platforms is to be retrevied.

.LINK
http://msdn.microsoft.com/en-us/library/cc146485.aspx
#>
Function Get-SCCMProgramSupportedPlatforms {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]$program
    )

    # SupportedOperatingSystems is a lazy property, so we have to explicitly retrieve it with Get()
    $program.Get() | Out-Null
    return $program.SupportedOperatingSystems
}

<#
.SYNOPSIS
Sets the list of supported platforms for a program.

.DESCRIPTION
Sets the list of supported platforms for a program.  The list is an array of SMS_OS_Details objects whose individual value
that have to match the list of available platforms reported by the server.  In order to find out the list of available
platforms, you can use Get-SCCMSupportedPlatforms.  You can then use New-SCCMProgramSupportedPlatform to build an object
for each desired platform.

.PARAMETER program
The program to be configured.

.PARAMETER platformList
A list of SMS_OS_Details objects to be stored in the program configuration.

.LINK
http://msdn.microsoft.com/en-us/library/cc146485.aspx
#>
Function Set-SCCMProgramSupportedPlatforms {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0)]$program,
        [parameter(Mandatory=$true, Position=1)]$platformList
    )

    $program.Get() | Out-Null
    $program.SupportedOperatingSystems = $platformList
    $program.Put() | Out-Null
}

<#
.SYNOPSIS
Creates a recurring SCCM interval schedule token.

.DESCRIPTION
Creates a recurring SCCM interval schedule token.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3 character site code.

.PARAMETER dayDuration
The number of days in the interval.

.PARAMETER daySpan
The number of days spanning the interval.  Range 0-31.

.PARAMETER hourDuration
The number of hours in the interval.  Range is 0-23.

.PARAMETER hourSpan
Number of hours spanning intervals.  Range is 0-23.

.PARAMETER isGmt
Determines whether the schedule time is based on GMT.

.PARAMETER minuteDuration
The number of minutes in the interval.  Range is 0-59.

.PARAMETER minuteSpan
Number of minutes spanning intervals.  Range is 0-59.

.PARAMETER startTime
The time and date when the interval will be available.

.LINK
http://msdn.microsoft.com/en-us/library/cc145924.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc146489.aspx
#>
Function New-SCCMRecurIntervalScheduleToken {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateScript( { $_ -gt 0 } )][parameter(Position=0)][int]$dayDuration = 0,
        [ValidateRange(0,31)][parameter(Position=1)][int]$daySpan = 0,
        [ValidateRange(0,23)][parameter(Position=2)][int]$hourDuration = 0,
        [ValidateRange(0,23)][parameter(Position=3)][int]$hourSpan = 0,
        [parameter(Position=4)][boolean]$isGmt = 0,
        [ValidateRange(0,59)][parameter(Position=5)][int]$minuteDuration = 0,
        [ValidateRange(0,59)][parameter(Position=6)][int]$minuteSpan = 0,
        [parameter(Mandatory=$true, Position=7)][DateTime]$startTime
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $scheduleToken = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ST_RecurInterval")).CreateInstance()   
    if($scheduleToken) {
        $scheduleToken.DayDuration = $dayDuration
        $scheduleToken.DaySpan = $daySpan
        $scheduleToken.HourDuration = $hourDuration
        $scheduleToken.HourSpan = $hourSpan
        $scheduleToken.IsGMT = $isGmt
        $scheduleToken.MinuteDuration = $minuteDuration
        $scheduleToken.MinuteSpan = $minuteSpan
        $scheduleToken.StartTime = (Convert-DateToSCCMDate $startTime)

        return $scheduleToken
    } else {
        Throw "Unable to create a new recurring interval schedule token"
    }
}

<#
.SYNOPSIS
Creates a non recurring SCCM interval schedule token.

.DESCRIPTION
Creates a non recurring SCCM interval schedule token.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3 character site code.

.PARAMETER dayDuration
The number of days in the interval.

.PARAMETER hourDuration
The number of hours in the interval.  Range is 0-23.

.PARAMETER isGmt
Determines whether the schedule time is based on GMT.

.PARAMETER minuteDuration
The number of minutes in the interval.  Range is 0-59.

.PARAMETER startTime
The time and date when the interval will be available.

.LINK
http://msdn.microsoft.com/en-us/library/cc145924.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc143487.aspx
#>
Function New-SCCMNonRecurringScheduleToken {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateScript( { $_ -gt 0 } )][parameter(Position=0)][int]$dayDuration = 0,
        [ValidateRange(0,23)][parameter(Position=1)][int]$hourDuration = 0,
        [parameter(Position=2)][boolean]$isGmt = 0,
        [ValidateRange(0,59)][parameter(Position=3)][int]$minuteDuration = 0,
        [parameter(Mandatory=$true, Position=4)][DateTime]$startTime
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $scheduleToken = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ST_NonRecurring")).CreateInstance()   
    if($scheduleToken) {
        $scheduleToken.DayDuration = $dayDuration
        $scheduleToken.HourDuration = $hourDuration
        $scheduleToken.IsGMT = $isGmt
        $scheduleToken.MinuteDuration = $minuteDuration
        $scheduleToken.StartTime = (Convert-DateToSCCMDate $startTime)

        return $scheduleToken
    } else {
        Throw "Unable to create a new non-recurring interval schedule token"
    }
}

<#
.SYNOPSIS
Creates a recurring schedule token that happens on specific days of the month at specific monthly intervals.

.DESCRIPTION
Creates a recurring schedule token that happens on specific days of the month at specific monthly intervals.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3 character site code.

.PARAMETER dayDuration
The number of days in the interval.

.PARAMETER forNumberOfMonths
Number of months in the interval.  Range is 1-12.

.PARAMETER hourDuration
The number of hours in the interval.  Range is 0-23.

.PARAMETER isGmt
Determines whether the schedule time is based on GMT.

.PARAMETER minuteDuration
The number of minutes in the interval.  Range is 0-59.

.PARAMETER monthDay
Day of the month when the event happens.  Range is 0-31 with 0 indicating the last day of the month.

.PARAMETER startTime
The time and date when the interval will be available.

.LINK
http://msdn.microsoft.com/en-us/library/cc145924.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc146724.aspx
#>
Function New-SCCMRecurMonthlyByDateScheduleToken {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateScript( { $_ -gt 0 } )][parameter(Position=0)][int]$dayDuration = 0,
        [ValidateRange(1,12)][parameter(Position=1)][int]$forNumberofMonths = 0,
        [ValidateRange(0,23)][parameter(Position=2)][int]$hourDuration = 0,
        [parameter(Position=3)][boolean]$isGmt = 0,
        [ValidateRange(0,59)][parameter(Position=4)][int]$minuteDuration = 0,
        [ValidateRange(0,31)][parameter(Position=5)][int]$monthDay = 0,
        [parameter(Mandatory=$true, Position=6)][DateTime]$startTime
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $scheduleToken = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ST_RecurMonthlyByDate")).CreateInstance()   
    if($scheduleToken) {
        $scheduleToken.DayDuration = $dayDuration
        $scheduleToken.ForNumberofMonths = $forNumberofMonths
        $scheduleToken.HourDuration = $hourDuration
        $scheduleToken.IsGMT = $isGmt
        $scheduleToken.MinuteDuration = $minuteDuration
        $scheduleToken.MonthDay = $monthDay
        $scheduleToken.StartTime = (Convert-DateToSCCMDate $startTime)

        return $scheduleToken
    } else {
        Throw "Unable to create a new monthly-by-date recurring interval schedule token"
    }    
}

<#
.SYNOPSIS
Creates a recurring schedule token that happens on specific days of the week at specific monthly intervals.

.DESCRIPTION
Creates a recurring schedule token that happens on specific days of the week at specific monthly intervals.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3 character site code.

.PARAMETER day
The day of the month when the event happens.

.PARAMETER dayDuration
The number of days in the interval.

.PARAMETER forNumberOfMonths
Number of months in the interval.  Range is 1-12.

.PARAMETER hourDuration
The number of hours in the interval.  Range is 0-23.

.PARAMETER isGmt
Determines whether the schedule time is based on GMT.

.PARAMETER minuteDuration
The number of minutes in the interval.  Range is 0-59.

.PARAMETER startTime
The time and date when the interval will be available.

.PARAMETER weekOrder
The week of the month when the event happens.  Range is 0-4.

0 - LAST
1 - FIRST
2 - SECOND
3 - THIRD
4 - FOURTH

.LINK
http://msdn.microsoft.com/en-us/library/cc144566.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc146724.aspx
#>
Function New-SCCMRecurMonthlyByWeekdayScheduleToken {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateRange(1,7)][parameter(Position=0)][int]$day = 0,
        [ValidateScript( { $_ -gt 0 } )][parameter(Position=1)][int]$dayDuration = 0,
        [ValidateRange(1,12)][parameter(Position=2)][int]$forNumberofMonths = 0,
        [ValidateRange(0,23)][parameter(Position=3)][int]$hourDuration = 0,
        [parameter(Position=4)][boolean]$isGmt = 0,
        [ValidateRange(0,59)][parameter(Position=5)][int]$minuteDuration = 0,
        [parameter(Mandatory=$true, Position=6)][DateTime]$startTime,
        [ValidateRange(0,4)][parameter(Position=7)][int]$weekOrder = 0
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $scheduleToken = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ST_RecurMonthlyByWeekday")).CreateInstance()   
    if($scheduleToken) {
        $scheduleToken.Day = $day
        $scheduleToken.DayDuration = $dayDuration
        $scheduleToken.ForNumberofMonths = $forNumberofMonths
        $scheduleToken.HourDuration = $hourDuration
        $scheduleToken.IsGMT = $isGmt
        $scheduleToken.MinuteDuration = $minuteDuration
        $scheduleToken.StartTime = (Convert-DateToSCCMDate $startTime)
        $scheduleToken.WeekOrder = $weekOrder

        return $scheduleToken
    } else {
        Throw "Unable to create a new monthly-by-weekday recurring interval schedule token"
    }        
}

<#
.SYNOPSIS
Creates a recurring schedule token that happens on a weekly interval.

.DESCRIPTION
Creates a recurring schedule token that happens on a weekly interval.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3 character site code.

.PARAMETER day
The day of the week when the event happens.

.PARAMETER dayDuration
The number of days in the interval.

.PARAMETER forNumberOfWeeks
Number of weeks for recurrence.  Range is 1-4.

.PARAMETER hourDuration
The number of hours in the interval.  Range is 0-23.

.PARAMETER isGmt
Determines whether the schedule time is based on GMT.

.PARAMETER minuteDuration
The number of minutes in the interval.  Range is 0-59.

.PARAMETER startTime
The time and date when the interval will be available.

.LINK
http://msdn.microsoft.com/en-us/library/cc146527.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc146724.aspx
#>
Function New-SCCMRecurWeeklyScheduleToken {
    [CmdletBinding()]
    param(
        [string]$siteProvider,
        [string]$siteCode,
        [ValidateRange(1,7)][parameter(Position=0)][int]$day = 0,
        [ValidateScript( { $_ -gt 0 } )][parameter(Position=1)][int]$dayDuration = 0,
        [ValidateRange(1,4)][parameter(Position=2)][int]$forNumberofWeeks = 0,
        [ValidateRange(0,23)][parameter(Position=3)][int]$hourDuration = 0,
        [parameter(Position=4)][boolean]$isGmt = 0,
        [ValidateRange(0,59)][parameter(Position=5)][int]$minuteDuration = 0,
        [parameter(Mandatory=$true, Position=6)][DateTime]$startTime
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $scheduleToken = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ST_RecurWeekly")).CreateInstance()   
    if($scheduleToken) {
        $scheduleToken.Day = $day
        $scheduleToken.DayDuration = $dayDuration
        $scheduleToken.ForNumberofWeeks = $forNumberofWeeks
        $scheduleToken.HourDuration = $hourDuration
        $scheduleToken.IsGMT = $isGmt
        $scheduleToken.MinuteDuration = $minuteDuration
        $scheduleToken.StartTime = (Convert-DateToSCCMDate $startTime)

        return $scheduleToken
    } else {
        Throw "Unable to create a new weekly recurring interval schedule token"
    }        
}

<#
.SYNOPSIS
Returns an SCCM folder object.

.DESCRIPTION
Returns an SCCM folder object.  Currently, the objects returned can only be of Type 2 or 3 (Package folders and Advertisement folders, respectively).
If no folder name or ID are specified, this function returns all folders on the site.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER folderName
The name of the folder to be returned.  If more than one folder is found with the same name, an array containing all matches is returned.

.PARAMETER folderNodeId
Unique folder ID of the folder to be returned.

.LINK
http://msdn.microsoft.com/en-us/library/cc145264.aspx
#>
Function Get-SCCMFolder {
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
        [parameter(Position=0)]
        [parameter(ParameterSetName="name")]
        [string]$folderName,
        [parameter(ParameterSetName="id")]
        [ValidateScript( { $_ -gt 0 } )]
        [int]$folderNodeId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($folderName) {
         return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.Name -like $folderName) -and (@(2,3) -contains $_.ObjectType) }
    } elseif($folderNodeId) {
         return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.ContainerNodeID -eq $folderNodeId) -and (@(2,3) -contains $_.ObjectType) }
    } else {
         return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { @(2,3) -contains $_.ObjectType }
    }
}

<#
.SYNOPSIS
Creates a new SCCM folder.

.DESCRIPTION
Creates a new SCCM folder.  Currently, the supported folder types are Package and Advertisement folders.  If there is already a folder with the specified
name under the same parent, an exception will be raised.  

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER folderName
The name of the new folder.

.PARAMETER folderType
The type of folder.  Allowed values are 2 and 3 (Package and Advertisement, respectively).

.PARAMETER parentFolderNodeID
ID of the parent folder.  If you specify 0, the folder will be placed under the root node for that folder type.

.LINK
http://msdn.microsoft.com/en-us/library/cc145264.aspx
#>
Function New-SCCMFolder {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateNotNull()][string]$folderName,
        [parameter(Mandatory=$true, Position=1)][ValidateRange(2,3)][int]$folderType,
        [parameter(Mandatory=$true, Position=2)][ValidateScript( { $_ -ge 0 } )][int]$parentFolderNodeId=0
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }    

    # We need to find out if there's another folder with that name under the same parent.  If there is, we throw an exception and don't create the folder.
    $allFolders = Get-SCCMFolder -siteProvider $siteProvider -siteCode $siteCode
    foreach($folder in $allFolders) {
        if( ($folder.Name -eq $folderName) -and ($folder.ParentContainerNodeID -eq $parentFolderNodeId) ) {
            Throw "There is already a folder named $folderName under the same parent folder with ID $parentFolderNodeId"
        }
    }

    # If we made it down here is because we need to go ahead with the folder creation.
    $folder = ([wmiclass]("\\$siteProvider\ROOT\sms\site_" + $siteCode + ":SMS_ObjectContainerNode")).CreateInstance()
    if($folder) {
        $folder.Name = $folderName
        $folder.ObjectType = $folderType
        $folder.ParentContainerNodeID = $parentFolderNodeId
        $folder.Put() | Out-Null

        return $folder
    } else {
        Throw "There was a problem creating the folder"
    }
}

<#
.SYNOPSIS
Saves a folder back into the SCCM database.

.DESCRIPTION
This function is used to save direct property changes to folders back to the SCCM database.

.PARAMETER folder
The folder object to be put back into the database.
#>
Function Save-SCCMFolder {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$folder
    )

    $folder.Put() | Out-Null
}

<#
.SYNOPSIS
Delete an SCCM folder.

.DESCRIPTION
Deletes and SCCM folder.  If you delete a parent folder without deleting its children, SCCM will by default assign the topmost
child folder the parent folder ID of 0.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER folderNodeId
The unique node ID for the folder being deleted.

.LINK
http://msdn.microsoft.com/en-us/library/cc145264.aspx
#>
Function Remove-SCCMFolder {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateScript( { $_ -gt 0 } )][ValidateNotNull()][int]$folderNodeId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $folder = Get-SCCMFolder -siteProvider $siteProvider -siteCode $siteCode -folderNodeId $folderNodeId
    if($folder) {
        $folder.Delete() | Out-Null
    }
}

<#
.SYNOPSIS
Moves an SCCM Folder.

.DESCRIPTION
Moves an SCCM Folder.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER folderNodeId
The unique node ID for the folder being moved.

.PARAMETER newParentFolderNodeId
The unique node ID of the new parent folder.

.LINK
http://msdn.microsoft.com/en-us/library/cc145264.aspx
#>
Function Move-SCCMFolder {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateScript( { $_ -gt 0 } )][ValidateNotNull()][int]$folderNodeId,
        [parameter(Mandatory=$true, Position=1)][ValidateScript( { $_ -ge 0 } )][ValidateNotNull()][int]$newParentFolderNodeId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    # We need to make sure the parent actually exists, if it doesn't, we're moving the folder to the root folder for that object type.
    if($newParentFolderNodeId -ne 0) {
        $parentFolder = Get-SCCMFolder -siteProvider $siteProvider -siteCode $siteCode -folderNodeId $newParentFolderNodeId
        if(!$parentFolder) {
            $newParentFolderNodeId = 0
        }
    }

    $folder = Get-SCCMFolder -siteProvider $siteProvider -siteCode $siteCode -folderNodeId $folderNodeId
    if($folder) {
        $folder.ParentContainerNodeID = $newParentFolderNodeId
        $folder.Put() | Out-Null
        return $folder
    }    
}

<#
.SYNOPSIS
Moves objects between containers.

.DESCRIPTION
Moves objects between containers.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER packageId
The instance key of the object to be moved.

.PARAMETER targetContainerNodeId
Node ID of the target container.  If it is 0, the object is moved back to the root container of objects of the same type.

.PARAMETER objectType
The type of object being moved.  Current supported values are 2 (packages) and 3 (advertisements).

.LINK
http://msdn.microsoft.com/en-us/library/cc144997.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc146279.aspx

.LINK
http://msdn.microsoft.com/en-us/library/cc145264.aspx
#>
Function Move-ObjectToContainer {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateLength(8,8)][string]$instanceKey,
        [parameter(Mandatory=$true, Position=1)][ValidateScript( { $_ -ge 0 } )][ValidateNotNull()][int]$targetContainerNodeId,
        [parameter(Mandatory=$true, Position=3)][ValidateRange(2,3)][int]$objectType
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $sourceContainerId = 0
    $sourceContainer = Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerItem" | Where { $_.InstanceKey -eq $instanceKey }
    if($sourceContainer) {
        # The object is in a folder other than the root folder.
        $sourceContainerId = $sourceContainer.ContainerNodeID
    }

    $containerClass = [WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ObjectContainerItem")
    $result = $containerClass.MoveMembers($instanceKey, $sourceContainerId, $targetContainerNodeId, $objectType)
    return $result.ReturnValue 
}

<#
.SYNOPSIS
Moves an SCCM Package to a folder.

.DESCRIPTION
Moves an SCCM Package to a folder.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER packageId
The ID of the package to be moved.

.PARAMETER targetFolderNodeId
The folder to which the package is to be moved.  If this value is 0, the package is removed from all folders.
#>
Function Move-SCCMPackageToFolder {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateLength(8,8)][string]$packageId,
        [parameter(Mandatory=$true, Position=1)][ValidateScript( { $_ -ge 0 } )][ValidateNotNull()][int]$targetFolderNodeId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $result = Move-ObjectToContainer -siteProvider $siteProvider -siteCode $siteCode -instanceKey $packageId -targetContainerNodeId $targetFolderNodeId -objectType 2
    if($result -ne 0) {
        Throw "There was a problem moving package with ID $packageId to folder with ID $folderNodeId"
    }
}

<#
.SYNOPSIS
Moves an SCCM advertisement to a folder.

.DESCRIPTION
Moves an SCCM advertisement to a folder.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER advertisementId
The ID of the advertisement to be moved.

.PARAMETER targetFolderNodeId
The folder to which the advertisement is to be moved.  If this value is 0, the advertisement is removed from all folders.
#>
Function Move-SCCMAdvertisementToFolder {
    [CmdletBinding()]
    param (
        [string]$siteProvider,
        [string]$siteCode,
        [parameter(Mandatory=$true, Position=0)][ValidateLength(8,8)][string]$advertisementId,
        [parameter(Mandatory=$true, Position=1)][ValidateScript( { $_ -ge 0 } )][ValidateNotNull()][int]$targetFolderNodeId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $result = Move-ObjectToContainer -siteProvider $siteProvider -siteCode $siteCode -instanceKey $advertisementId -targetContainerNodeId $targetFolderNodeId -objectType 3
    if($result -ne 0) {
        Throw "There was a problem moving advertisement with ID $advertisementId to folder with ID $folderNodeId"
    }
}

<#
.SYNOPSIS
Utility function to convert DMTF date strings into something readable and usable by PowerShell.

.DESCRIPTION
This function exists for convenience so users of the module do not have to try to figure out, if they are not familiar with it, ways
to convert DMTF date strings into something readable.
#>
Function Convert-SCCMDateToDate {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$date
    )

    [System.Management.ManagementDateTimeconverter]::ToDateTime($date);
}

<#
.SYNOPSIS
Utility function to convert PowerShell date strings into DMTF dates.

.DESCRIPTION
This function exists for convenience so users of the module do not have to try to figure out, if they are not familiar with it, ways
to convert PowerShell date strings into DMTF dates.

.PARAMETER date
The date string to be converted to a DMTF date.

.EXAMPLE
Convert-DateToSCCMDate $(Get-Date)

Description
-----------
This will convert the current date and time into a DMTF date string that SCCM understands.
#>
Function Convert-DateToSCCMDate {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$date
    )

    [System.Management.ManagementDateTimeconverter]::ToDMTFDateTime($date)
}

Set-Alias -Name "gsa" -Value Get-SCCMAdvertisement
Set-Alias -Name "gscol" -Value Get-SCCMCollection
Set-Alias -Name "gsc" -Value Get-SCCMComputer
Set-Alias -Name "gsdist" -Value Get-SCCMDistributionPoints
Set-Alias -Name "gsf" -Value Get-SCCMFolder
Set-Alias -Name "gspk" -Value Get-SCCMPackage
Set-Alias -Name "gspg" -Value Get-SCCMProgram

Export-ModuleMember Add-SCCMComputerToCollection
Export-ModuleMember Add-SCCMPackageToDistributionPoint
Export-ModuleMember Convert-DateToSCCMDate
Export-ModuleMember Convert-SCCMDateToDate
Export-ModuleMember Get-SCCMAdvertisement -Alias "gsa"
Export-ModuleMember Get-SCCMAdvertisementAssignedSchedule
Export-ModuleMember Get-SCCMAdvertisementsForCollection
Export-ModuleMember Get-SCCMAdvertisementsForComputer
Export-ModuleMember Get-SCCMAdvertisementsForPackage
Export-ModuleMember Get-SCCMAdvertisementStatusForComputer
Export-ModuleMember Get-SCCMClientAdvertisementScheduleId
Export-ModuleMember Get-SCCMClientAssignedSite
Export-ModuleMember Get-SCCMClientCacheSize
Export-ModuleMember Get-SCCMClientSoftwareDistributionHistory 
Export-ModuleMember Get-SCCMCollection -Alias "gscol"
Export-ModuleMember Get-SCCMCollectionMembers
Export-ModuleMember Get-SCCMCollectionRefreshSchedule
Export-ModuleMember Get-SCCMCollectionsForComputer
Export-ModuleMember Get-SCCMComputer -Alias "gsc"
Export-ModuleMember Get-SCCMComputerVariables
Export-ModuleMember Get-SCCMDistributionPoints -Alias "gsdist"
Export-ModuleMember Get-SCCMFolder -Alias "gsf"
Export-ModuleMember Get-SCCMMaintenanceWindows
Export-ModuleMember Get-SCCMMaintenanceWindowSchedules
Export-ModuleMember Get-SCCMPackage -Alias "gspk"
Export-ModuleMember Get-SCCMProgram -Alias "gspg"
Export-ModuleMember Get-SCCMProgramSupportedPlatforms
Export-ModuleMember Get-SCCMSupportedPlatforms
Export-ModuleMember Invoke-SCCMClientAction
Export-ModuleMember Invoke-SCCMClientSchedule
Export-ModuleMember Move-SCCMAdvertisementToFolder
Export-ModuleMember Move-SCCMFolder
Export-ModuleMember Move-SCCMPackageToFolder
Export-ModuleMember New-SCCMAdvertisement
Export-ModuleMember New-SCCMComputer
Export-ModuleMember New-SCCMComputerVariable
Export-ModuleMember New-SCCMFolder
Export-ModuleMember New-SCCMNonRecurringScheduleToken
Export-ModuleMember New-SCCMPackage
Export-ModuleMember New-SCCMProgram
Export-ModuleMember New-SCCMRecurIntervalScheduleToken
Export-ModuleMember New-SCCMRecurMonthlyByDateScheduleToken
Export-ModuleMember New-SCCMRecurMonthlyByWeekdayScheduleToken
Export-ModuleMember New-SCCMRecurWeeklyScheduleToken
Export-ModuleMember New-SCCMStaticCollection
Export-ModuleMember New-SCCMSupportedPlatform
Export-ModuleMember Remove-SCCMAdvertisement
Export-ModuleMember Remove-SCCMCollection
Export-ModuleMember Remove-SCCMComputer
Export-ModuleMember Remove-SCCMComputerFromCollection
Export-ModuleMember Remove-SCCMFolder
Export-ModuleMember Remove-SCCMPackage
Export-ModuleMember Remove-SCCMPackageFromDistributionPoint
Export-ModuleMember Remove-SCCMProgram
Export-ModuleMember Save-SCCMAdvertisement
Export-ModuleMember Save-SCCMCollection
Export-ModuleMember Save-SCCMFolder
Export-ModuleMember Save-SCCMPackage
Export-ModuleMember Save-SCCMProgram
Export-ModuleMember Set-SCCMAdvertisementAssignedSchedule
Export-ModuleMember Set-SCCMClientAssignedSite
Export-ModuleMember Set-SCCMClientCacheSize
Export-ModuleMember Set-SCCMCollectionRefreshSchedule
Export-ModuleMember Set-SCCMComputerVariables
Export-ModuleMember Set-SCCMProgramSupportedPlatforms