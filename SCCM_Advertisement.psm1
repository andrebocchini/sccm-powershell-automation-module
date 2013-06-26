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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $advertisementName,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateLength(8,8)]
        [string]
        $collectionId,
        [parameter(Mandatory=$true, Position=2)]
        [ValidateLength(8,8)]
        [string]
        $packageId,
        [parameter(Mandatory=$true, Position=3)]
        [ValidateNotNull()]
        [string]
        $programName
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
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $advertisement
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $advertisementId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    $advertisement = Get-SCCMAdvertisement -siteProvider $siteProvider -siteCode $siteCode -advertisementId $advertisementId
    if($advertisement) {
        $advertisement.Delete()
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
        [string]
        $siteProvider,
        [parameter(ParameterSetName="name")]
        [parameter(ParameterSetName="default")]
        [parameter(ParameterSetName="id")]
        [string]
        $siteCode,
        [parameter(ParameterSetName="name", Position=0, ValueFromPipeline=$true)]
        [ValidateNotNull()]
        [string]
        $advertisementName,
        [parameter(Position=1)]
        [parameter(ParameterSetName="id")]
        [ValidateLength(8,8)]
        [ValidateNotNull()]
        [string]
        $advertisementId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($advertisementName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * from SMS_Advertisement WHERE AdvertisementName like '$advertisementName%'"
    } elseif($advertisementId) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Advertisement" -filter "AdvertisementID='$advertisementId'"
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
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateNotNull()]
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $packageId
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Advertisement" -Filter "PackageID='$packageId'"
}
