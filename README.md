# GetSiteCollectionInventory
This is a Console App that can be used to generate an inventory report of all lists / libraries, items in the lists / libraries. 

It uses SharePoint CSOM, which means this does NOT have to be run on the SharePoint server. 

It can be run locally from a machine that is on the network, and just needs to be able to access the SharePoint on-prem site via web browser.

It needs to be run using credentials that has Site Collection Admin rights to the SharePoint on-prem site.

This app specifically captures the following info in the inventory report (.CSV file)
Site Title, Site URL, List Title, List Item Count, List URL, List Type, Item Id, Item Type, Item Title, Created, Created by, Modified, Modified by, Item URL

Nuget packages referenced:
Latest versions of the following as of September 5th 2018
- Microsoft.SharePointOnline.CSOM
- SharePointPnPCoreOnline

Replace the settings in the appSettings section in the app.config file with your SharePoint on-premise related info.