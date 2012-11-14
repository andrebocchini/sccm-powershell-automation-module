About
=====

This Powershell module contains a collection of functions gathered from an assortment of scripts I use to automate SCCM 2007 management.  I realized that every time I needed to automate something I ended up digging through a pile of old scripts to cut and paste code into a new one, so eventually I decided I needed to aggregate all of the bits of code I had created into a single script.  This module does not encompass a large part of client management functions in SCCM, but it can perform a number of common operations such as creating and deleting new computer records, packages, advertisements and manipulating collection membership rules.  Most of this code came about because of my own need to automate software distribution, but a lot of it can be used for other types of tasks.

These are the functions currently present in the module:

Site Functions

    Add-SCCMComputerToCollection
    Add-SCCMPackageToDistributionPoint
    Get-SCCMAdvertisement
    Get-SCCMAdvertisementAssignedSchedule
    Get-SCCMAdvertisementsForCollection
    Get-SCCMAdvertisementsForComputer
    Get-SCCMAdvertisementsForPackage
    Get-SCCMAdvertisementStatusForComputer
    Get-SCCMCollection
    Get-SCCMCollectionMembers
    Get-SCCMCollectionsForComputer
    Get-SCCMComputer
    Get-SCCMComputerVariables
    Get-SCCMDistributionPoints
    Get-SCCMFolder
    Get-SCCMMaintenanceWindows
    Get-SCCMMaintenanceWindowSchedules
    Get-SCCMPackage
    Get-SCCMProgram
    Get-SCCMProgramSupportedPlatforms
    Get-SCCMSupportedPlatforms
    Move-SCCMAdvertisementToFolder
    Move-SCCMFolder
    Move-SCCMPackageToFolder
    New-SCCMAdvertisement
    New-SCCMComputer
    New-SCCMFolder
    New-SCCMNonRecurringScheduleToken
    New-SCCMPackage
    New-SCCMProgram
    New-SCCMRecurIntervalScheduleToken
    New-SCCMRecurMonthlyByDateScheduleToken
    New-SCCMRecurMonthlyByWeekdayScheduleToken
    New-SCCMRecurWeeklyScheduleToken
    New-SCCMStaticCollection
    New-SCCMSupportedPlatform
    Remove-SCCMAdvertisement
    Remove-SCCMCollection
    Remove-SCCMComputer
    Remove-SCCMComputerFromCollection
    Remove-SCCMComputerVariable
    Remove-SCCMFolder
    Remove-SCCMPackage
    Remove-SCCMPackageFromDistributionPoint
    Remove-SCCMProgram
    Save-SCCMAdvertisement
    Save-SCCMCollection
    Save-SCCMFolder
    Save-SCCMPackage
    Save-SCCMProgram
    Set-SCCMAdvertisementAssignedSchedule
    Set-SCCMComputerVariable
    Set-SCCMProgramSupportedPlatforms

Client Functions

    Get-SCCMClientAdvertisementScheduleId
    Get-SCCMClientAssignedSite
    Get-SCCMClientCacheSize
    Get-SCCMClientSoftwareDistributionHistory 
    Invoke-SCCMClientAction
    Invoke-SCCMClientSchedule
    Set-SCCMClientAssignedSite
    Set-SCCMClientCacheSize

Utility Functions

    Convert-DateToSCCMDate
    Convert-SCCMDateToDate
   
This code has only been tested with SCCM 2007.  Please conduct your own independent testing before trusting this code in a production environment.

Installation
============

You will need to have Powershell V3 installed.  If you don't already have it, you can download it here:

http://www.microsoft.com/en-us/download/details.aspx?id=34595

Make sure the following directory exists and is part of your PSModulePath:

    %userprofile%\Documents\WindowsPowershell\Modules

If you run the Install.ps1 script, it should copy the files to the right place for you.  If the script fails to copy the files properly, you can do so manually by creating the directory:

    %userprofile%\Documents\WindowsPowershell\Modules\SCCM

and copying the following files into it:

    SCCM.psm1
    SCCM.psd1
    SCCM_Formats.ps1xml

After the files are in place, you should be able to run:

    Import-Module SCCM
    Get-Help SCCM

Usage
=====

Use the following line at the top of your scripts:
    
    Import-Module SCCM

If the import is successful, you should be able to use all of the module's exported functions.  To see a list of available functions, use:

    Get-Help SCCM

Examples
========

Here are a couple of examples of things you can automate with this module.  Take, for example, a simple workflow where an SCCM admin has to perform the following tasks:

1. Create a new package
2. Create an installation program for the new package
3. Distribute it to distribution points
4. Create a collection in order to test this package
5. Add a computer to the test collection
6. Advertise the program to the test collection
7. Instruct the test computer to retrieve new machine policies

Here's how this can be accomplished:

1. Create a new package

        $newPackage = New-SCCMPackage `
                            -packageName $packageName `
                            -packageDescription $packageDescription `
                            -packageVersion $packageVersion `
                            -packageManufacturer $packageManufacturer `
                            -packageLanguage $packageLanguage `
                            -packageSource $packageSourcePath

    The package will be created with default settings, but you can customize it further if you wish.  Look at this link for some more information about package flags and settings http://msdn.microsoft.com/en-us/library/cc144959.aspx.  If you do make changes to the package, make sure to use Save-SCCMPackage when you're finished.

2. Create an installation program for the new package

        $newProgram = New-SCCMProgram `
                            -packageId $newPackage.PackageID `
                            -programName $programName `
                            -programCommandLine $programCommandLine

    The program will be created with default settings, but you can customize it further if you wish.  Look at this link for some more information about program flags and settings http://msdn.microsoft.com/en-us/library/cc144361.aspx. If you do make changes to the program, make sure to use Save-SCCMProgram when you're finished.

3. Distribute it to distribution points
    
        $distributionPoints = Get-SCCMDistributionPoints
    
        Add-SCCMPackageToDistributionPoint `
                            -packageId $newPackage.packageID `
                            -distributionPointList $distributionPoints

4. Create a collection in order to test the package

        $newCollection = New-SCCMStaticCollection `
                            -collectionName $newCollectionName `
                            -parentCollectionId $parentCollectionId

5. Add a computer to the test collection

        $testComputer = Get-SCCMComputer -computerName $testComputerName

        Add-SCCMComputerToCollection `
                            -resourceId $testComputer.ResourceID
                            -collectionId $testCollection.CollectionID

6. Advertise the program to the test collection

        $newAdvertisement = New-SCCMAdvertisement `
                            -advertisementName $advertisementName `
                            -collectionId $newCollection.CollectionID `
                            -packageId $newPackage.PackageID `
                            -programName $newProgram.ProgramName

    The advertisement will be created with default settings, but you can customize it further if you wish.  Look at this link for some more information about advertisement flags and settings http://msdn.microsoft.com/en-us/library/cc146108.aspx.  If you do make changes to the advertisement, make sure to use Save-SCCMAdvertisement when you're finished.

7. Instruct the test computer to retrieve new machine policies

        Invoke-SCCMClientAction `
                            -computerName $testComputerName
                            -scheduleId "MachinePolicyRetrievalEvalCycle"