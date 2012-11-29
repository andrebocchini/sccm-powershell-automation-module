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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [int]
        $resourceId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [int]
        $resourceId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [int]
        $resourceId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $collectionName,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateLength(8,8)]
        [string]
        $parentCollectionId,
        [string]
        $collectionComment
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
    )

    Set-Variable refreshTypeManual -option Constant -value 1
    Set-Variable refreshTypeAuto -option Constant -value 2

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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateRange(1,2)]
        [int]
        $refreshType = 1,
        [parameter(Position=2)]
        [ValidateScript( { !(!$_ -and $refreshType -eq 2) } )]
        $refreshSchedule
    )
    
    Set-Variable refreshTypeManual -option Constant -value 1
    Set-Variable refreshTypeAuto -option Constant -value 2

    if($refreshType -eq $refreshTypeAuto) {
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
    if($refreshType -eq $refreshTypeAuto) {
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
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $collection
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if($collection) {
        $collection.Delete()
    } else {
        Throw "Unable to retrieve and delete collection with ID $collectionId"
    }
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
        [string]
        $siteProvider,
        [parameter(ParameterSetName="name")]
        [parameter(ParameterSetName="default")]
        [parameter(ParameterSetName="id")]
        [string]
        $siteCode,
        [parameter(ParameterSetName="name", Position=0, ValueFromPipeline=$true)]
        [string]
        $collectionName,
        [parameter(ParameterSetName="id", Position=1)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [int]
        $resourceId
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
Retrieves collection settings.

.DESCRIPTION
Retrieves collection settings.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
The ID of the collection whose settings are to be retrieved.

.LINK
http://msdn.microsoft.com/en-us/library/cc145320.aspx
#>
Function Get-SCCMCollectionSettings {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collectionSettings = Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_CollectionSettings" | where {$_.CollectionID -eq $collectionId}
    if($collectionSettings) {
        $collectionSettings.Get() | Out-Null
    }
    return $collectionSettings
} 

<#
.SYNOPSIS
Creates collection settings objects.

.DESCRIPTION
Creates collection settings objects.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
The ID of the collection for which settings are to be created.

.LINK
http://msdn.microsoft.com/en-us/library/cc145320.aspx
#>
Function New-SCCMCollectionSettings {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collection) {
        Throw "Unable to retrive collection with ID $collectionId"
    }

    $collectionSettings = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_CollectionSettings")).CreateInstance()
    if($collectionSettings) {
        $collectionSettings.CollectionID = $collection.CollectionID
    }
    return $collectionSettings
}

<#
.SYNOPSIS
Sets variables on an SCCM collection.

.DESCRIPTION
Takes in information about a specific site, along with a collection ID and an array of variables, and assigns those variables to the collection.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER resourceId
ID of the target collection.

.PARAMETER variableList
An array of variables to be assigned to the collection.  If you pass it an empty array, it will clear all variables on the collection.

.LINK
http://msdn.microsoft.com/en-us/library/cc145320.aspx
#>
Function Set-SCCMCollectionVariables {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateNotNull()]
        $variableList
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collection) {
        Throw "Unable to retrieve collection with ID $collectionId"
    }

    $collectionSettings = Get-SCCMCollectionSettings -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collectionSettings) {
        $collectionSettings = New-SCCMCollectionSettings -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    }
    $collectionSettings.CollectionVariables = $variableList
    Save-SCCMCollectionSettings $collectionSettings
}

<#
.SYNOPSIS
Returns an array of all collection variables for a specific collection.

.DESCRIPTION
Returns an array of all collection variables for a specific collection.

.PARAMETER siteProvider
Name of the site provider.

.PARAMETER siteCode
3-character site code.

.PARAMETER collectionId
ID of the target collection.

.LINK
http://msdn.microsoft.com/en-us/library/cc146201.aspx
#>
Function Get-SCCMCollectionVariables {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collection = Get-SCCMCollection -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collection) {
        Throw "Unable to retrieve collection with ID $collectionId"
    }

    $collectionSettings = Get-SCCMCollectionSettings -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if($collectionSettings) {
        return $collectionSettings.CollectionVariables
    }

}

<#
.SYNOPSIS
Creates a new comllection variable.

.DESCRIPTION
Creates a new collection variable.

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
http://msdn.microsoft.com/en-us/library/cc146201.aspx
#>
Function New-SCCMCollectionVariable {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $variableName,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $variableValue,
        [parameter(Mandatory=$true, Position=0)]
        [bool]
        $isMasked
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collectionVariable = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_CollectionVariable")).CreateInstance()
    if($collectionVariable) {
        $collectionVariable.IsMasked = $isMasked
        $collectionVariable.Name = $variableName
        $collectionVariable.Value = $variableValue
    }
    return $collectionVariable
}

<#
.SYNOPSIS
Saves an object containing a collection's machine settings back into the database.

.DESCRIPTION
Saves an object containing a collection's machine settings back into the database.

.PARAMETER collectionSettings
The object to be saved back into the database.

.LINK
http://msdn.microsoft.com/en-us/library/cc145320.aspx
#>
Function Save-SCCMCollectionSettings {
    [CmdletBinding()]
    param (        
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $collectionSettings
    )

    $collectionSettings.Put() | Out-Null
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId
    )    

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collectionSettings = Get-SCCMCollectionSettings -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        $maintenanceWindow
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
Creates a new maintenance window object.

.DESCRIPTION
Creates a new maintenance window object.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER windowName
A descriptive name for the maintenance window.

.PARAMETER windowDescription
A description for the maintenance window.

.PARAMETER windowSchedules
Schedule token object representing the maintenance window schedule.

.PARAMETER windowRecurrenceType
An integer indicating the recurrence type of the window.  Allowed values are:

1 NONE
2 DAILY
3 WEEKLY
4 MONTHLYBYWEEKDAY
5 MONTHLYBYDATE

If you define a recurrence type that does not match the type of schedule tokens pased to this function,
you will encounter erros when you attempt to assign this maintenance window to a collection.  For example,
if you choose recurrence type 1, you can only use non recurring schedule tokens.

.PARAMETER windowIsEnabled
A boolean value indicating whether this maintenance window is enabled.

.PARAMETER windowType
Allowed values are:

1 GENERAL. General maintenance window. 
5 OSD. Operating system deployment task sequence maintenance window.
 
.PARAMETER startTime
The date and time when this window will become available.d

.LINK
http://msdn.microsoft.com/en-us/library/cc143300.aspx
#>
Function New-SCCMMaintenanceWindow {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $windowName,
        [parameter(Position=1)]
        [string]
        $windowDescription,
        [parameter(Mandatory=$true, Position=2)]
        [ValidateNotNull()]
        $windowSchedules,
        [parameter(Mandatory=$true, Position=3)]
        [ValidateRange(1,5)]
        [int]
        $windowRecurrenceType,
        [parameter(Mandatory=$true, Position=4)]
        [boolean]
        $windowsIsEnabled,
        [parameter(Mandatory=$true, Position=5)]
        [ValidateScript( { ($_ -eq 1) -or ($_ -eq 5) } )]
        [int]
        $windowType,
        [parameter(Mandatory=$true, Position=6)]
        $startTime
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }    

    # We have to convert the schedule tokens to schedule strings
    if($windowSchedules) {
        $scheduleMethod = [WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ScheduleMethods")
        $result = $scheduleMethod.WriteToString($windowSchedules)
        $windowScheduleStrings = $result.StringData
    }

    $maintenanceWindow = ([WMIClass]("\\$siteProvider\root\sms\site_" + "$siteCode" + ":SMS_ServiceWindow")).CreateInstance()
    if($maintenanceWindow) {
        $maintenanceWindow.Name = $windowName
        $maintenanceWindow.Description = $windowDescription
        $maintenanceWindow.ServiceWindowSchedules = $windowScheduleStrings
        $maintenanceWindow.RecurrenceType = $windowRecurrenceType
        $maintenanceWindow.IsEnabled = $windowIsEnabled
        $maintenanceWindow.ServiceWindowType = $windowType
        $maintenanceWindow.StartTime = Convert-DateToSCCMDate $startTime        
    }
    return $maintenanceWindow
}

<#
.SYNOPSIS
Assigns an array of maintenance window objects to a collection.

.DESCRIPTION
Assigns an array of maintenance window objects to a collection.

.PARAMETER siteProvider
The name of the site provider.

.PARAMETER siteCode
The 3-character site code.

.PARAMETER collectionId
The ID of the collection that will received the maintenance windows.

.PARAMETER maintenanceWindows
An array of maintenance window objects.  If this is an empty array, all maintenance windows will be removed from
the collection.

.LINK
http://msdn.microsoft.com/en-us/library/cc143300.aspx
#>
Function Set-SCCMCollectionMaintenanceWindows {
    [CmdletBinding()]
    param (
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $collectionId,
        [parameter(Mandatory=$true, Position=1)]
        [AllowEmptyCollection()]
        [ValidateNotNull()]
        [array]
        $maintenanceWindows
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $collectionSettings = Get-SCCMCollectionSettings -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    if(!$collectionSettings) {
        # If the collection has never had any settings, we have to create some fresh ones.
        $collectionSettings = New-SCCMCollectionSettings -siteProvider $siteProvider -siteCode $siteCode -collectionId $collectionId
    }

    if($collectionSettings) {
        $collectionSettings.ServiceWindows = $maintenanceWindows
        Save-SCCMCollectionSettings $collectionSettings
    } else {
        Throw "Unable to retrieve collection settings from collection with ID $collectionId"
    }
}