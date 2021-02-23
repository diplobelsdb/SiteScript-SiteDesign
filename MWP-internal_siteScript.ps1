# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
#https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs
# Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"
# get all site designs
# Get-SPOSiteDesign


$internal_site_script = @'
  {
    "$schema": "schema.json",    
    "actions": [            
     {
  		"verb": "setSiteExternalSharingCapability",
  		"capability": "Disabled"
	 }        
    ],
    "bindata": { },
    "version": 1
}
'@

# add SiteScript
Add-SPOSiteScript  -Title "Internal MWP"  -Content $internal_site_script  -Description "Disabling external sharing"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "Internal MWP"} | Select -First 1 Id
$guid.Id.ToString()  


# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Internal MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $internal_site_script -Version 1


# ----------------------------------------------

# display sitescript content 
# $sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Internal MWP"} | Select -First 1 Id
# Get-SPOSiteScript $sitescriptID.Id 

# Get scripts
# Get-SPOSiteScript | where {$_.Title -EQ "Internal MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}


