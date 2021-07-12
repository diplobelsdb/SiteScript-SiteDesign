# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
#https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs


Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"

# Get all SiteScripts
get-spositescript


$posten_en_site_script = @'
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
                "title": "01.Communication - Public diplomacy"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
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
                "title": "02.Premises, IT and logistics"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
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
                "title": "03.Protocol and administration"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
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
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      },{
        "verb": "createSPList",
        "listName": "05Finance",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "05.Financial policy"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      } ,{
        "verb": "createSPList",
        "listName": "06Consul",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "06.Consular affairs and assistance to Belgian nationals"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      },{
        "verb": "createSPList",
        "listName": "07ForeignAffairs",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "07.Foreign Policy"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      }  ,{
        "verb": "createSPList",
        "listName": "08GlobalMngt",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "08.General management"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      }  ,{
        "verb": "createSPList",
        "listName": "09Security",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "09.Security"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      },{
        "verb": "createSPList",
        "listName": "Documents",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "10.Documents Teams"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      } ,{
        "verb": "createSPList",
        "listName": "98Pics",
        "templateType": 109,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "98.Pictures"
            }
        ]
      },{
        "verb": "createSPList",
        "listName": "99Archives",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "99.Archives"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      } ,{
        "verb": "createSPList",
        "listName": "Agenda",
        "templateType": 106,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "Calendar"
            }
        ]
      }   ,{
        "verb": "createSPList",
        "listName": "Archive_ShareDoc",
        "templateType": 101,
        "addNavLink":true,
        "subactions": [
            {
                "verb": "setTitle",
                "title": "Archives ShareDoc"
            }
            ,{
                "verb": "addSPView",
                "name": "ArchiveInfo",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString",
                "ArchiveCode",
                "ToBeArchived"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Default"
            },
            {
                "verb": "addSPView",
                "name": "No folders",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": false,
                "scope": "Recursive"
            },
            {
                "verb": "addSPView",
                "name": "Versions",
                "viewFields":
                [
                "DocIcon",
                "LinkFilename",
                "Modified",
                "Editor",
                "_UIVersionString"
                ],
                "query": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>",
                "rowLimit": 30,
                "isPaged": true,
                "makeDefault": true,
                "scope": "Default"
            }
        ]
      }          
    ],
    "bindata": { },
    "version": 1
}
'@

# add SiteScript
Add-SPOSiteScript  -Title "Posten EN MWP"  -Content $posten_en_site_script -Description "Adding 10 Document Libraries for 'Posten EN', a Calendar and a picture library"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "Posten EN MWP"} | Select -First 1 Id
$guid.Id.ToString()  


# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Posten EN MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $posten_en_site_script -Version 1


# ---------
# display sitescript content 
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Posten EN MWP"} | Select -First 1 Id
Get-SPOSiteScript $sitescriptID.Id 


# remove scripts
# Get-SPOSiteScript | where {$_.Title -EQ "Posten EN MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}

