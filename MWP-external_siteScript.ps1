# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
# https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs


Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"


$external_site_script = @'
{
    "$schema": "schema.json",    
    "actions": [            
     {
  		"verb": "setSiteExternalSharingCapability",
  		"capability": "ExternalUserSharingOnly"
	 }              
    ],
    "bindata": { },
    "version": 1
}
'@

# add SPView
#https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema#addspview

# Add SiteScript
Add-SPOSiteScript  -Title "External MWP"  -Content $external_site_script  -Description "Enable external sharing"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "External MWP"} | Select -First 1 Id
$guid.Id.ToString() 

# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "External MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $external_site_script -Version 1



# ----------------------------------------------

# display sitescript content 
# $sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "External MWP"} | Select -First 1 Id
# Get-SPOSiteScript $sitescriptID.Id 

# Get scripts
# Get-SPOSiteScript | where {$_.Title -EQ "External MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}



