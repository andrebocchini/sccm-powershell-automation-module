About
=====

This Powershell module contains a collection of functions gathered from an assortment of scripts I use to automate SCCM management.  It does not encompass a large part of client management functions in SCCM, but it can perform a number of common operations such as creating new computer records, deleting them, and manipulating collection membership rules.  It has only been tested with SCCM 2007.  Please conduct your own independent testing before trusting this code in a production environment.

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