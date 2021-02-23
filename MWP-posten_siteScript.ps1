# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
#https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs


Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"

# Get all SiteScripts
get-spositescript


$posten_site_script = @'
{
    "$schema": "schema.json",
    "actions": [  
      {
        "verb": "createSPList",
        "listName": "01Communication_PublicDiplomacy",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "01. Communication Public Diplomacy"
            }
        ]
      },           
      {
        "verb": "createSPList",
        "listName": "02Premises_IT_Logistics",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "02. Premises IT Logistics"
            }
        ]
      },{
        "verb": "createSPList",
        "listName": "03Protocol_Administration",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "03. Protocol Administration"
            }
        ]
      },{
        "verb": "createSPList",
        "listName": "04Human_Resources",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "04. Human Resources"
            }
        ]
      }        
    ],
    "bindata": { },
    "version": 1
}
'@

# add SiteScript
Add-SPOSiteScript  -Title "Posten MWP"  -Content $posten_site_script  -Description "Adding 4 Document Libraries for 'Posten'"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "Posten MWP"} | Select -First 1 Id
$guid.Id.ToString()  


# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Posten MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $posten_site_script -Version 1


# ---------
# display sitescript content 
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Posten MWP"} | Select -First 1 Id
Get-SPOSiteScript $sitescriptID.Id 


# remove scripts
# Get-SPOSiteScript | where {$_.Title -EQ "Posten MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}

