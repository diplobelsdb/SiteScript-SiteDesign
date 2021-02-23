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


# add SiteDesign
# Modern Team site - 64
Add-SPOSiteDesign -Title "MWP – internal Team site" -WebTemplate "64" -SiteScripts @($guid_internal, $guid_global)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Team site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_global)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# Posten
Add-SPOSiteDesign -Title "MWP – internal Posten site" -WebTemplate "64" -SiteScripts @($guid_internal, $guid_posten, $guid_global)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Posten site" -WebTemplate "64" -SiteScripts @($guid_external, $guid_posten, $guid_global)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 


# Communication site - 68
Add-SPOSiteDesign -Title "MWP – internal Communication site" -WebTemplate "68" -SiteScripts @($guid_internal, $guid_global)  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 
Add-SPOSiteDesign -Title "MWP – external Communication site" -WebTemplate "68" -SiteScripts @($guid_external, $guid_global)  -Description "Enabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

# Add-SPOSiteDesign -Title "MWP – internal Communication site" -WebTemplate "68" -SiteScripts $guid.Id.ToString()  -Description "Disabling external sharing, adding theme, adding custom columns for archiving, setting regional settings, customizing navigation" 

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

