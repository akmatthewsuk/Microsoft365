<#
   .NOTES
    ===========================================================================
    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
    EITHER EXPRESSED OR IMPLIED,  INCLUDING BUT NOT LIMITED TO THE IMPLIED
    WARRANTIES OF MERCHANTBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
    ===========================================================================
    Created by:    Andrew Matthews
    Filename:      Get-MessageTracking.ps1
    Documentation: None
    Execution Tested on: Exchange 2016 installed
    Requires:     Exchange PowerShell module
    Versions:
    1.0 - Initial Release 21-Oct-2022
    ===========================================================================

    .SYNOPSIS
    Report on message tracking for a sender/recipient pair

    .DESCRIPTION
    Queries Exchange On-Prem and outputs message tracking for a sender/recipient pair
    
    .PARAMETER Sender
    Specifies the sender

    .PARAMETER Recipient
    Specifies the recipient

    .PARAMETER TimePeriod
    specifies the time period to search back from the current date/time

    .EXAMPLE
    C:\PS>Get-MessageTracking.ps1 -Sender test@testdomain.com -Recipient test2@testdomain.com
    
    .EXAMPLE
    C:\PS>Get-MessageTracking.ps1 -Sender test@testdomain.com -Recipient test2@testdomain.com -TimePeriod 6
#>
Param (
    [Parameter(Mandatory=$True, HelpMessage="Specify the sender")][String]$Sender,
    [Parameter(Mandatory=$True, HelpMessage="Specify the recipient")][String]$Recipient,
    [Parameter(Mandatory=$false, HelpMessage="Specify the time period in hours to search (default 1 hour)")][Int]$TimePeriod = 1
)

Write-Host "Obtaining Exchange Server List"
try{
    $EXServers = Get-ExchangeServer -ErrorAction Stop | Sort Name
} catch {
    $ErrorMessage = $_.Exception.Message
    Write-Host "Failed to load the list of Exchange servers" -ForegroundColor Red
    Write-host $ErrorMessage -ForegroundColor Red
    Exit
}

if($Sender -like "*@*") {
    #Valid recipient
} else {
    #invalid recipient
    Write-Host "The specified sender ($($Sender)) is not a valid email address" -ForegroundColor Red
    Exit
}
if($Recipient -like "*@*") {
    #Valid recipient
} else {
    #invalid recipient
    Write-Host "The specified recipient ($($Recipient)) is not a valid email address" -ForegroundColor Red
    Exit
}

if($TimePeriod -eq 1) {
    Write-Host "Querying messages from $($Sender) to $($Recipient) in the last $TimePeriod hour"
} else {
    Write-Host "Querying messages from $($Sender) to $($Recipient) in the last $TimePeriod hours"
}

Foreach($EXServer in $EXServers) {
    Write-Host "Querying $($EXServer.Name)" -foregroundcolor blue
    try{
        Get-MessageTrackingLog -Sender $Sender -Recipients $Recipient -server $EXServer.name -Start (Get-date).AddHours(-$TimePeriod) -ErrorAction Stop | Where-object {$_.EventId -notlike "HA*" } | Select TimeStamp, EventID, MessageSubject, clienthostname,Serverhostname, ConnectorId | ft -autosize
    } catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Failed to query message tracking" -ForegroundColor Red
        Write-host $ErrorMessage -ForegroundColor Red
    }
 }