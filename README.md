About
=====

This Powershell module contains a collection of functions gathered from an assortment of scripts I use to automate SCCM 2007 management.  I realized that every time I needed to automate something I ended up digging through a pile of old scripts to cut and paste code into a new one, so eventually I decided I needed to aggregate all of the bits of code I had created into a single script.  This module does not encompass a large part of client management functions in SCCM, but it can perform a number of common operations such as creating and deleting new computer records, packages, advertisements and manipulating collection membership rules.  Most of this code came about because of my own need to automate software distribution, but a lot of it can be used for other types of tasks.

These are the functions currently present in the module:

Site Functions

    New-SCCMComputer
    Remove-SCCMComputer
    Get-SCCMComputer
    Add-SCCMComputerToCollection
    Remove-SCCMComputerFromCollection
    New-SCCMStaticCollection
    Remove-SCCMCollection
    Get-SCCMCollection
    Get-SCCMCollectionMembers
    Get-SCCMCollectionsForComputer
    New-SCCMAdvertisement
    Save-SCCMAdvertisement
    Remove-SCCMAdvertisement
    Get-SCCMAdvertisement
    Get-SCCMAdvertisementsForCollection
    Get-SCCMAdvertisementsForComputer
    Get-SCCMAdvertisementsForPackage
    Get-SCCMAdvertisementStatusForComputer
    Set-SCCMComputerVariable
    Get-SCCMComputerVariables
    Remove-SCCMComputerVariable
    New-SCCMPackage
    Save-SCCMPackage
    Remove-SCCMPackage
    Get-SCCMPackage
    New-SCCMProgram
    Save-SCCMProgram
    Remove-SCCMProgram
    Get-SCCMProgram
    Add-SCCMPackageToDistributionPoint
    Remove-SCCMPackageFromDistributionPoint
    Get-SCCMDistributionPoints

Client Functions

    Invoke-SCCMClientAction
    Invoke-SCCMClientSchedule
    Get-SCCMClientSoftwareDistributionHistory 
    Get-SCCMClientAdvertisementScheduleId
    Get-SCCMClientAssignedSite
    Set-SCCMClientAssignedSite
    Get-SCCMClientCacheSize
    Set-SCCMClientCacheSize

Utility Functions

    Convert-SCCMDateToDate
    Convert-DateToSCCMDate

This code has only been tested with SCCM 2007.  Please conduct your own independent testing before trusting this code in a production environment.

Installation
============

Run the Install.ps1 script provided with the module.  Alternatively, create the directory:

    %userprofile%\Documents\WindowsPowershell\Modules\SCCM

After creating the directory, copy the SCCM.psm1 file into it.

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

        $newPackage = New-SCCMPackage -siteProvider $siteProvider -siteCode $siteCode `
                            -packageName $packageName `
                            -packageDescription $packageDescription `
                            -packageVersion $packageVersion `
                            -packageManufacturer $packageManufacturer `
                            -packageLanguage $packageLanguage `
                            -packageSource $packageSourcePath

The package will be created with default settings, but you can customize it further if you wish.  Look at this link for some help http://msdn.microsoft.com/en-us/library/cc144959.aspx.  If you do make changes to the package, make sure to use Save-SCCMPackage when you're finished.

2. Create an installation program for the new package

        $newProgram = New-SCCMProgram `
                            -siteProvider $siteProvider `
                            -siteCode $siteCode `
                            -packageId $newPackage.PackageID `
                            -programName $programName `
                            -programCommandLine $programCommandLine

The program will be created with default settings, but you can customize it further if you wish.  Look at this link for some help http://msdn.microsoft.com/en-us/library/cc144361.aspx. If you do make changes to the program, make sure to use Save-SCCMProgram when you're finished.

3. Distribute it to distribution points
    
        $distributionPoints = Get-SCCMDistributionPoints -siteProvider $siteProvider -siteCode $siteCode
    
        Add-SCCMPackageToDistributionPoint `
                            -siteProvider $siteProvider `
                            -siteCode $siteCode `
                            -packageId $newPackage.packageID `
                            -distributionPointList $distributionPoints

4. Create a collection in order to test the package

        $newCollection = New-SCCMStaticCollection `
                            -siteProvider $siteProvider `
                            -siteCode $siteCode `
                            -collectionName $newCollectionName `
                            -parentCollectionId $parentCollectionId

5. Add a computer to the test collection

        Add-SCCMComputerToCollection `
                            -siteProvider $siteProvider
                            -siteCode $siteCode
                            -computerName $testComputerName
                            -collectionName $testCollection.Name

6. Advertise the program to the test collection

        $newAdvertisement = New-SCCMAdvertisement `
                            -siteProvider $siteProvider `
                            -siteCode $siteCode `
                            -advertisementName $advertisementName `
                            -collectionId $newCollection.CollectionID `
                            -packageId $newPackage.PackageID `
                            -programName $newProgram.ProgramName

The advertisement will be created with default settings, but you can customize it further if you wish.  Look at this link for some help http://msdn.microsoft.com/en-us/library/cc146108.aspx.  If you do make changes to the advertisement, make sure to use Save-SCCMAdvertisement when you're finished.

7. Instruct the test computer to retrieve new machine policies

        Invoke-SCCMClientAction `
                            -computerName $testComputerName
                            -scheduleId "MachinePolicyRetrievalEvalCycle"