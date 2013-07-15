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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [string]
        $packageName,
        [parameter(Mandatory=$true, Position=1)]
        [AllowEmptyString()]
        [string]
        $packageDescription = "",
        [parameter(Mandatory=$true, Position=2)]
        [AllowEmptyString()]
        [string]
        $packageVersion = "",
        [parameter(Mandatory=$true, Position=3)]
        [AllowEmptyString()]
        [string]
        $packageManufacturer = "",
        [AllowEmptyString()]
        [string]
        $packageLanguage = "",
        [AllowEmptyString()]
        [string]
        $packageSource = ""   
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
        $newPackageId = $($newPackageIdTokens[1]).Replace("`"", "")

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
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $package
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

    $package = Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId
    if($package) {
        $package.Delete() | Out-Null
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
        $packageName,
        [parameter(ParameterSetName="id")]
        [parameter(Position=1)]
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

    if($packageName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "SELECT * FROM SMS_Package WHERE Name like '$packageName%'"
    } elseif($packageId) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Package" -filter "PackageID='$packageId'"
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [string]
        [ValidateLength(8,8)]
        $packageId,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateNotNull()]
        [string]
        $programName,
        [parameter(Mandatory=$true, Position=2)]
        [ValidateNotNull()]
        [string]        
        $programCommandLine
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

            return Get-SCCMProgram -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId -programName $programName
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
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $program
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $packageId,
        [parameter(Mandatory=$true, Position=1)]
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

    $program = Get-SCCMProgram -siteProvider $siteProvider -siteCode $siteCode $packageId $programName
    if($program) {
        return $program.Delete()
    } else {
        Throw "Invalid package ID $packageId or program named `"$programName`""
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Position=0)]
        [ValidateLength(8,8)]
        [string]
        $packageId,
        [parameter(Position=1)]
        [string]
        $programName
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if($packageId -and $programName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "SELECT * FROM SMS_Program WHERE PackageID='$packageId' AND ProgramName='$programName'"
    } elseif($packageId -and !$programName) {
        return Get-WMIObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Class "SMS_Program" -Filter "PackageID='$packageId'"
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]        
        [string]
        $packageId,
        [parameter(Mandatory=$true, Position=1)]
        $distributionPointList
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
            $newDistributionPoint.Put() | Out-Null
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateLength(8,8)]
        [string]
        $packageId,
        [parameter(Mandatory=$true, Position=1)]
        $distributionPointList
    )

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    if(Get-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode -packageId $packageId) {
        foreach($distributionPoint in $distributionPointList) {
            $distributionPointToBeDeleted = Get-WmiObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "SELECT * FROM SMS_DistributionPoint WHERE ServerNALPath='$($distributionPoint.NALPath)' AND PackageID='$packageId'"
            if($distributionPointToBeDeleted) {
                $distributionPointToBeDeleted.Delete()
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
        [string]
        $siteProvider,
        [string]
        $siteCode
    )    

    if(!($PSBoundParameters) -or !($PSBoundParameters.siteProvider)) {
        $siteProvider = Get-SCCMSiteProvider
    }
    if(!($PSBoundParameters) -or !($PSBoundParameters.siteCode)) {
        $siteCode = Get-SCCMSiteCode
    }

    return Get-WmiObject -ComputerName $siteProvider -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_SystemResourceList WHERE RoleName='SMS Distribution Point' AND SiteCode='$siteCode'"
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
        [string]
        $siteProvider,
        [string]
        $siteCode
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
        [string]
        $siteProvider,
        [string]
        $siteCode,
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        [string]
        $name,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateNotNull()]
        [string]
        $maxVersion,
        [parameter(Mandatory=$true, Position=2)]
        [ValidateNotNull()]
        [string]
        $minVersion,
        [parameter(Mandatory=$true, Position=3)]
        [ValidateNotNull()]
        [string]
        $platform
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
        [parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $program
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
        [parameter(Mandatory=$true, Position=0)]
        [ValidateNotNull()]
        $program,
        [parameter(Mandatory=$true, Position=1)]
        $platformList
    )

    $program.Get() | Out-Null
    $program.SupportedOperatingSystems = $platformList
    $program.Put() | Out-Null
}
