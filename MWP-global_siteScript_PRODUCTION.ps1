# Connect-SPOService -url "https://diplomatiebel-admin.sharepoint.com"

 

# get all site scripts
# Get-SPOSiteScript

 

        
$global_site_script = @'
 {
    "$schema": "schema.json",
    "actions": [            
    {
      "verb": "createSiteColumn",
      "group": "MWP",
      "fieldType": "Boolean",
      "internalName": "ToBeArchived",
      "displayName": "To be archived",
      "isRequired": false,
      "id": "68597476-4a78-40ab-bc90-95d8066c1359"
    }, 
    {
        "verb": "createSiteColumnXml",
        "schemaXml": "<Field Type=\"TaxonomyFieldType\" DisplayName=\"Archive code\" ShowField=\"Term1033\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" ID=\"{95e7a709-00b2-4fdf-9bcf-b9b5fc846c60}\" StaticName=\"ArchiveCode\" Name=\"ArchiveCode\" Group=\"MWP\"><Default /><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value xmlns:q1=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q1:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">8710b318-ea48-4423-a308-0e87359dff93</Value></Property><Property><Name>GroupId</Name></Property><Property><Name>TermSetId</Name><Value xmlns:q2=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q2:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">eca26591-3e39-4461-87f0-273b620e3239</Value></Property><Property><Name>AnchorId</Name><Value xmlns:q3=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q3:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">00000000-0000-0000-0000-000000000000</Value></Property><Property><Name>UserCreated</Name><Value xmlns:q4=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q4:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>Open</Name><Value xmlns:q5=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q5:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>IsPathRendered</Name><Value xmlns:q7=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q7:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>IsKeyword</Name><Value xmlns:q8=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q8:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>TargetTemplate</Name></Property><Property><Name>CreateValuesInEditForm</Name><Value xmlns:q9=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q9:boolean\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">false</Value></Property><Property><Name>FilterAssemblyStrongName</Name><Value xmlns:q10=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q10:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">Microsoft.SharePoint.Taxonomy, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Value></Property><Property><Name>FilterClassName</Name><Value xmlns:q11=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q11:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">Microsoft.SharePoint.Taxonomy.TaxonomyField</Value></Property><Property><Name>FilterMethodName</Name><Value xmlns:q12=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q12:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">GetFilteringHtml</Value></Property><Property><Name>FilterJavascriptProperty</Name><Value xmlns:q13=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q13:string\" xmlns:p4=\"http://www.w3.org/2001/XMLSchema-instance\">FilteringJavascript</Value></Property></ArrayOfProperty></Customization></Field>"
    } ,
    {
        "verb": "createContentType",
        "name": "Document",
        "description": "Existing content type",
        "parentName": "Item",
        "hidden": false,
        "subactions": [                         
            {
            "verb": "addSiteColumn",
            "internalName":"ArchiveCode"
            },
            {
            "verb": "addSiteColumn",
            "internalName":"ToBeArchived"
            }
        ]
    },
    {
        "verb": "createSPList",
        "listName": "Documents",
        "templateType": 101,  
        "addNavLink":true,
        "subactions": [
            {
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
              "query": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>",
              "rowLimit": 30,
              "isPaged": true,
              "makeDefault": true,
              "scope": "Default"
            }
        ]
    },
    {
        "verb": "applyTheme",
        "themeName": "MWP theme"
    },
    {
      "verb": "setRegionalSettings",
      "timeZone": 3, /* Brussels, Copenhagen, Madrid, Paris */
      "locale": 2057, /* 2057 = English(UK) --  2067 = Dutch Belgium  --  2060 = French Belgium */
      "sortOrder": 25, /* General */
      "hourFormat": "24"
    } ,
    {
        "verb": "removeNavLink",
        "displayName": "Pages",
        "isWebRelative": true
    }
    ],
    "bindata": { },
    "version": 1
}
'@


# add SiteScript
Add-SPOSiteScript  -Title "Global MWP"  -Content $global_site_script  -Description "Adding theme, adding custom columns for archiving, setting regional settings, customizing navigation"
$guid = Get-SPOSiteScript | where {$_.Title -EQ "Global MWP"} | Select -First 1 Id
$guid.Id.ToString()  

 


# update a previously uploaded Site Script
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Global MWP"} | Select -First 1 Id
Set-SPOSiteScript -Identity $sitescriptID.Id  -Content $global_site_script -Version 1

 

# ----------------------------------------------------
# display sitescript content 
$sitescriptID = Get-SPOSiteScript | where {$_.Title -EQ "Global MWP"} | Select -First 1 Id
Get-SPOSiteScript $sitescriptID.Id 

  

# remove scripts
 Get-SPOSiteScript | where {$_.Title -EQ "Internal MWP"} | ForEach-Object {Remove-SPOSiteScript -Identity $_.Id.ToString()}