# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
#https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs


Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"
# get all site designs
Get-SPOSiteDesign

# get guids of SiteScripts
$guid_global = (Get-SPOSiteScript | where {$_.Title -EQ "Global MWP"} | Select -First 1 Id).Id.Guid
$guid_external = (Get-SPOSiteScript | where {$_.Title -EQ "External MWP"} | Select -First 1 Id).Id.Guid
$guid_internal = (Get-SPOSiteScript | where {$_.Title -EQ "Internal MWP"} | Select -First 1 Id).Id.Guid
$guid_posten = (Get-SPOSiteScript | where {$_.Title -EQ "Posten MWP"} | Select -First 1 Id).Id.Guid
$guid_posten_FR = (Get-SPOSiteScript | where {$_.Title -EQ "Posten FR MWP"} | Select -First 1 Id).Id.Guid
$guid_posten_EN = (Get-SPOSiteScript | where {$_.Title -EQ "Posten EN MWP"} | Select -First 1 Id).Id.Guid
$guid_EXTRANET = (Get-SPOSiteScript | where {$_.Title -EQ "EXTRANET MWP"} | Select -First 1 Id).Id.Guid

# export content sitescript to notepad
$siteScriptTitle = 'TestListView'
$id = get-spositescript | Where-Object Title -eq $siteScriptTitle | select Id
get-spositescript -Identity $id.Id | select content | Out-File -FilePath "C:\data\SiteDesign\MWP-$($siteScriptTitle)_siteScript.ps1"

# update existing sitescript - use PnP
Connect-PnPOnline -Url https://diplobelsdb-admin.sharepoint.com/ -UseWebLogin

$siteScriptTitle = 'TestListView'
$id = Get-PnPSiteScript | Where-Object Title -eq $siteScriptTitle | select Id
Set-PnPSiteScript -Identity $id.Id -Description 'Test adding listview'



# add SiteDesign
# Modern Team site - 64
Add-SPOSiteDesign -Title "MWP – internal Team site" -WebTemplate "64" -SiteScripts @($guid_internal, $guid_global)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Team site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_global)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# Posten
Add-SPOSiteDesign -Title "MWP – internal Posten site" -WebTemplate "64" -SiteScripts @($guid_internal, $guid_posten, $guid_global)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Posten site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_posten, $guid_global)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# Posten FR
Add-SPOSiteDesign -Title "MWP – internal Posten FR site" -WebTemplate "64" -SiteScripts @($guid_internal, $guid_global, $guid_posten_FR)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Posten FR site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_global, $guid_posten_FR)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# Posten EN
Add-SPOSiteDesign -Title "MWP – internal Posten EN site" -WebTemplate "64" -SiteScripts @($guid_internal, $guid_global, $guid_posten_EN)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Posten EN site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_global, $guid_posten_EN)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 


# Communication site - 68
Add-SPOSiteDesign -Title "MWP – internal Communication site" -WebTemplate "68" -SiteScripts @($guid_internal, $guid_global)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Communication site" -WebTemplate "68" -SiteScripts @($guid_external, $guid_global)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# EXTRANET site - 64
Add-SPOSiteDesign -Title "MWP – EXTRANET site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_global, $guid_EXTRANET)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# ---------
# display sitescript content 
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Global MWP"} | Select -First 1 Id
Get-SPOSiteScript $sitescriptID.Id 

# $siteDesignIdentity = get-spositedesign | where {$_.Title -EQ "MWP – external Team site"} | Select -First 1 Id
# $siteDesignIdentity = get-spositedesign | where {$_.Title -EQ "MWP – external Communication site"} | Select -First 1 Id 
# Remove-SPOSiteDesign -Identity $siteDesignIdentity.Id

# display SiteDesign properties
get-spositedesign | where {$_.Title -EQ "MWP – external Communication site"} | Select Id, Title, Description, SiteScriptIds, WebTemplate | format-list

# update existing SiteDesign with new SiteScripts
$guidSiteDesign = (get-spositedesign | where {$_.Title -EQ "MWP – external Communication site"} | select id).id.guid
$guidSiteDesign = (get-spositedesign | where {$_.Title -EQ "MWP – external Team site"} | select id).id.guid
Set-SPOSiteDesign -Identity $guidSiteDesign -SiteScripts @($guid_external, $guid_global)

$guidSiteDesign = (get-spositedesign | where {$_.Title -EQ "MWP – internal Communication site"} | select id).id.guid
$guidSiteDesign = (get-spositedesign | where {$_.Title -EQ "MWP – internal Team site"} | select id).id.guid
Set-SPOSiteDesign -Identity $guidSiteDesign -SiteScripts @($guid_internal, $guid_global)

