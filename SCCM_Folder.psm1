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
        [string]
        $siteProvider,
        [parameter(ParameterSetName="name")]
        [parameter(ParameterSetName="default")]
        [parameter(ParameterSetName="id")]
        [string]
        $siteCode,
        [parameter(Position=0)]
        [parameter(ParameterSetName="name")]
        [string]
        $folderName,
        [parameter(ParameterSetName="id")]
        [ValidateScript( { $_ -gt 0 } )]
        [int]
        $folderNodeId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $folderName,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateRange(2,3)]
        [int]
        $folderType,
        [parameter(Mandatory=$true, Position=2)]
        [ValidateScript( { $_ -ge 0 } )]
        [int]
        $parentFolderNodeId=0
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
        if( ($folder.Name -eq $folderName) -and ($folder.ParentContainerNodeID -eq $parentFolderNodeId) -and ($folder.ObjectType -eq $folderType) ) {
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
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $folder
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [ValidateNotNull()]
        [int]
        $folderNodeId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateScript( { $_ -gt 0 } )]
        [ValidateNotNull()]
        [int]
        $folderNodeId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateScript( { $_ -ge 0 } )]
        [ValidateNotNull()][int]
        $newParentFolderNodeId
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $instanceKey,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateScript( { $_ -ge 0 } )]
        [ValidateNotNull()]
        [int]
        $targetContainerNodeId,
        [parameter(Mandatory=$true, Position=3)]
        [ValidateRange(2,3)]
        [int]
        $objectType
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $sourceContainerId = 0
    $sourceContainer = Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerItem Where InstanceKey='$instanceKey'"
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $packageId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateScript( { $_ -ge 0 } )]
        [ValidateNotNull()]
        [int]
        $targetFolderNodeId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $result = Move-ObjectToContainer -siteProvider $siteProvider -siteCode $siteCode -instanceKey $packageId -targetContainerNodeId $targetFolderNodeId -objectType 2
    if($result -ne 0) {
        Throw "There was a problem moving package with ID $packageId to folder with ID $targetFolderNodeId"
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $advertisementId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateScript( { $_ -ge 0 } )]
        [ValidateNotNull()]
        [int]
        $targetFolderNodeId
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
