# https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json
#https://laurakokkarinen.com/the-ultimate-guide-to-sharepoint-site-designs-and-site-scripts/#limiting-who-can-use-site-designs


Connect-SPOService -url "https://diplobelsdb-admin.sharepoint.com"

# Get all SiteScripts
get-spositescript


$posten_fr_site_script = @'
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
                "title": "01.Communication - Public Diplomacy"
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
                "title": "02.Bâtiments, IT et Logistique"
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
                "title": "03.Protocole et administration"
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
                "title": "04.Resources humaines"
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
                "title": "05.Gestion financière"
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
                "title": "06.Affaires consulaires et assistance aux Belges"
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
                "title": "07.Politique étrangère"
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
                "title": "08.Gestion générale"
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
                "title": "09.Sécurité"
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
                "title": "Documents Teams"
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
                "title": "98.Photos"
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
                "title": "Agenda"
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
Add-SPOSiteScript  -Title "Posten FR MWP"  -Content $posten_fr_site_script -Description "Adding 10 Document Libraries for 'Posten FR', a Calendar and a picture library"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "Posten FR MWP"} | Select -First 1 Id
$guid.Id.ToString()  


# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Posten FR MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $posten_fr_site_script -Version 1


# ---------
# display sitescript content 
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Posten FR MWP"} | Select -First 1 Id
Get-SPOSiteScript $sitescriptID.Id 


# remove scripts
# Get-SPOSiteScript | where {$_.Title -EQ "Posten FR MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}

