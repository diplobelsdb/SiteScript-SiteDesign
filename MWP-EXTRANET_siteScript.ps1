# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
#https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs


# Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"

# Get all SiteScripts
get-spositescript


$EXTRANET_site_script = @'
{
    "$schema": "schema.json",
    "actions": [  
       {
        "verb": "removeNavLink",
        "displayName": "Conversations",
        "isWebRelative": true
        } 
        , {
        "verb": "removeNavLink",
        "displayName": "Notebook",
        "isWebRelative": true
        }     
    ],
    "bindata": { },
    "version": 1
}
'@

# add SiteScript
Add-SPOSiteScript  -Title "EXTRANET MWP"  -Content $EXTRANET_site_script  -Description "Removing navigation link 'Conversations' and 'Notebook'"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "EXTRANET MWP"} | Select -First 1 Id
$guid.Id.ToString()  


# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "EXTRANET MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $EXTRANET_site_script -Version 1


# ---------
# display sitescript content 
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "EXTRANET MWP"} | Select -First 1 Id
Get-SPOSiteScript $sitescriptID.Id 


# remove scripts
# Get-SPOSiteScript | where {$_.Title -EQ "EXTRANET MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}

