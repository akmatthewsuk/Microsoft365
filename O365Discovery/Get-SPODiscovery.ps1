Write-Host "Getting SharePoint Online Settings"
$OutputPath = "c:\Temp\SPOOutput"
if((test-path -PathType Container -Path $OutputPath) -ne $True){
    Write-Host "Creating output path $($OutputPath)"
    New-Item -ItemType Directory -Path $OutputPath
}

$SPOOutput_TenantSettings = Join-Path -Path $OutputPath -ChildPath "SPO_TenantSettings.txt"
$SPOOutput_SiteDesign = Join-Path -Path $OutputPath -ChildPath "SPO_SiteDesigns.csv"
$SPOOutput_SiteTheme = Join-Path -Path $OutputPath -ChildPath "SPO_SiteTheme.csv"
$SPOOutput_SiteScripts = Join-Path -Path $OutputPath -ChildPath "SPO_SiteScripts.txt"
$SPOOutput_OrgAssets = Join-Path -Path $OutputPath -ChildPath "SPO_OrgAssets.csv"

Get-SPOTenant > $SPOOutput_TenantSettings
Get-SPOSiteDesign | Export-Csv -NoTypeInformation -Path $SPOOutput_SiteDesign



$SPOSiteThemeOutput = New-Object System.Collections.ArrayList
$SPOSiteThemes = Get-SPOTheme
If(($SPOSiteThemes).count -eq 0) {
    $SPOSiteThemeOutput.Add("No SharePoint Site Themes found") | Out-Null
} else {
    #Add the columns
    $SpoThemeHeaderLine="Name"
    foreach($Key in ($SPOSiteThemes[0].Palette.Keys | Sort-Object)){
        $SpoThemeHeaderLine = $SpoThemeHeaderLine + ",$($Key)"
    }
    $SPOSiteThemeOutput.Add($SpoThemeHeaderLine) | Out-Null
    ForEach($SPOTheme in $SPOSiteThemes) {
        #Add the SPO Theme Name
        $SpoThemeLine = "$($SPOTheme.Name)"
        #Get the key value pairs
        foreach($SPOThemeKeyValuePair in ($SPOTheme.Palette.GetEnumerator() | Sort-Object)) {
            $SpoThemeLine = $SpoThemeLine + ",$($SPOThemeKeyValuePair.Value)"
        }
        $SPOSiteThemeOutput.Add($SpoThemeLine) | Out-Null
    }
}
Add-Content -Path $SPOOutput_SiteTheme -Value $SPOSiteThemeOutput

$SPOSiteScriptOutput = New-Object System.Collections.ArrayList
$SPOSiteScripts = Get-SPOSiteScript
If(($SPOSiteScripts).count -eq 0) {
    $SPOSiteScriptOutput.Add("No SharePoint Site Scripts found") | Out-Null
} else {
    ForEach($SPOSiteScript in $SPOSiteScripts){
        
        $SPOSiteScriptOutput.Add("Title: $($SPOSiteScript.Title)") | Out-Null
        $SPOSiteScriptOutput.Add("Description: $($SPOSiteScript.Description)") | Out-Null
        $SPOSiteScriptOutput.Add("ID: $($SPOSiteScript.ID.guid)") | Out-Null
        $SPOSiteScriptOutput.Add("Version: $($SPOSiteScript.Version)") | Out-Null
        #Get the SPO SiteThem Content
        $GetSPOScriptContent = Get-SPOSiteScript -Identity "$($SPOSiteScript.ID.guid)"
        $SPOSiteScriptOutput.Add("Content") | Out-Null
        $SPOSiteScriptOutput.Add("$($GetSPOScriptContent)") | Out-Null
        $SPOSiteScriptOutput.Add("---------------------------------------") | Out-Null
    }

}
Add-Content -Path $SPOOutput_SiteScripts -Value $SPOSiteScriptOutput

Get-SPOOrgAssetsLibrary | Select-Object DisplayName, LibraryUrl, ListId, OrgAssetType | Export-CSV -NoTypeInformation -Path $SPOOutput_OrgAssets