<#
   .NOTES
    ===========================================================================
    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
    EITHER EXPRESSED OR IMPLIED,  INCLUDING BUT NOT LIMITED TO THE IMPLIED
    WARRANTIES OF MERCHANTBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
    ===========================================================================
    Created by:    Andrew Matthews
    Filename:      Get-Exchangequeues.ps1
    Documentation: None
    Execution Tested on: Exchange 2016 installed
    Requires:     Exchange PowerShell module
    Versions:
    1.0 - Initial Release 21-Oct-2022
    ===========================================================================

    .SYNOPSIS
    Report on Exchange server queue state

    .DESCRIPTION
    Queries exchange servers and outputs the state of queues

    .EXAMPLE
    C:\PS>Get-Exchangequeues.ps1
    
#>


Write-Host "Obtaining Exchange Server List"
try{
    $EXServers = Get-ExchangeServer -ErrorAction Stop | Sort Name
} catch {
    $ErrorMessage = $_.Exception.Message
    Write-Host "Failed to load the list of Exchange servers" -ForegroundColor Red
    Write-host $ErrorMessage -ForegroundColor Red
    Exit
}

Foreach($EXServer in $EXServers) {
    Write-Host "Querying $($EXServer.Name)" -foregroundcolor blue
    try{
        Get-Queue -server $EXServer.name | Where-Object {$_.MessageCount -gt 0} | Where-Object {$_.DeliveryType -ne "ShadowRedundancy"} | Sort Identity | Select-Object Identity, DeliveryType, MessageCount, NextHopDomain, LastError | Ft
    } catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Failed to query Exchange queues" -ForegroundColor Red
        Write-host $ErrorMessage -ForegroundColor Red
    }
 }