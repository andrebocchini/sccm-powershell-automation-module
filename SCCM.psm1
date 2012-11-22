<#
.SYNOPSIS
Command line interface for an assortment of SCCM operations.

.DESCRIPTION
The functions in this module provide a command line and scripting interface for automating 
management of SCCM environments.

.NOTES
File Name  : SCCM.psm1  
Author     : Andre Bocchini <andrebocchini@gmail.com>

.LINK
https://github.com/andrebocchini/SCCM-Powershell-Automation-Module
#>

<#
.SYNOPSIS
Attempts to discover the local computer's site provider.

.DESCRIPTION
When a user does not specify a site provider for a function that requires that information, we call this 
code to try to determine who the provider is automatically.  If we cannot find it, we throw an exception.
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
        Throw "Unable to determine site provider"
    } finally{
        $ErrorActionPreference = "Continue"
    }
}

<#
.SYNOPSIS
Attempts to discover the local computer's site code.

.DESCRIPTION
When a user does not specify a site code for a function that requires that information, we call this code 
to try to determine what the code is automatically.  If we cannot find it, we throw an exception.
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
         Throw "Unable to determine site code"
    } finally{
        $ErrorActionPreference = "Continue"
    }
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
Export-ModuleMember Get-SCCMCollectionVariables
Export-ModuleMember Get-SCCMComputer -Alias "gsc"
Export-ModuleMember Get-SCCMComputerVariables
Export-ModuleMember Get-SCCMDistributionPoints -Alias "gsdist"
Export-ModuleMember Get-SCCMFolder -Alias "gsf"
Export-ModuleMember Get-SCCMMaintenanceWindows
Export-ModuleMember Get-SCCMMaintenanceWindowSchedules
Export-ModuleMember Get-SCCMPackage -Alias "gspk"
Export-ModuleMember Get-SCCMProgram -Alias "gspg"
Export-ModuleMember Get-SCCMProgramSupportedPlatforms
Export-ModuleMember Get-SCCMSiteCode
Export-ModuleMember Get-SCCMSiteProvider
Export-ModuleMember Get-SCCMSupportedPlatforms
Export-ModuleMember Invoke-SCCMClientAction
Export-ModuleMember Invoke-SCCMClientSchedule
Export-ModuleMember Move-SCCMAdvertisementToFolder
Export-ModuleMember Move-SCCMFolder
Export-ModuleMember Move-SCCMPackageToFolder
Export-ModuleMember New-SCCMAdvertisement
Export-ModuleMember New-SCCMCollectionVariable
Export-ModuleMember New-SCCMComputer
Export-ModuleMember New-SCCMComputerVariable
Export-ModuleMember New-SCCMFolder
Export-ModuleMember New-SCCMMaintenanceWindow
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
Export-ModuleMember Set-SCCMCollectionMaintenanceWindows
Export-ModuleMember Set-SCCMCollectionRefreshSchedule
Export-ModuleMember Set-SCCMCollectionVariables
Export-ModuleMember Set-SCCMComputerVariables
Export-ModuleMember Set-SCCMProgramSupportedPlatforms