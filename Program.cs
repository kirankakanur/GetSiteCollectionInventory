using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.SharePoint.Client;


namespace GetSiteCollectionInventory
{
    class Program
    {
        static void Main(string[] args)
        {
            DoIt();
        }

        private static void DoIt()
        {
            // setup a stopwatch to compute the time it takes to completed the inventory
            var watch = System.Diagnostics.Stopwatch.StartNew();

            // get site collection url, site collection admin user name, password, domain, and file names from config file
            string siteUrl = ConfigurationManager.AppSettings["ClientContextUrl"];
            string userName = ConfigurationManager.AppSettings["ClientContextUsername"];
            string userPassword = ConfigurationManager.AppSettings["ClientContextPassword"];
            string domain = ConfigurationManager.AppSettings["ClientContextDomain"];
            string csvFilePath = ConfigurationManager.AppSettings["InventoryFilePath"];
            string csvLargeListsFilePath = ConfigurationManager.AppSettings["LargeListsFilePath"];

            // create StringBuilder object for storing site inventory info
            StringBuilder sbInvCSVFile = new StringBuilder();
            sbInvCSVFile.AppendLine("Site Title, Site URL, List Title, List Item Count, List URL, List Type, Item Id, Item Type, Item Title, Created, Created by, Modified, Modified by, Item URL");

            // create StringBuilder object for storing large lists info
            StringBuilder sbLargeListsCSVFile = new StringBuilder();
            sbLargeListsCSVFile.AppendLine("Site Title, Site URL, List Title, List Item Count, List URL, List Type");
            
            // these lists will NOT be inventoried
            string[] listExceptions = { "appdata", "Cache Profiles", "CacheData", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Content type service application error log", "Converted Forms", "Customized Reports", "Device Channels", "Form Templates", "fpdatasources", "Form Templates", "List Template Gallery", "Long Running Operation Status", "Master Page Gallery", "Phone Call Memo", "Reusable Content", "Reusable Images", "Shared Images", "TaxonomyHiddenList", "Theme Gallery", "User Information List", "Web Analytics Workflow History", "Web Part Gallery", "wfpub", "Workflow History", "Workflow Tasks", "Style Library", "List Template Libary", "Maintenance Log Library", "Solution Gallery", "Site Collection Images" };

            // items in generic lists with these IDs will get the DispForm.aspx?ID=<> as the default item url
            int[] genericListIds = { 100, 106, 120, 10010, 10012, 10014, 10020, 10022 };    

            // construct CAML Query to get ALL items from a given list, in descending Modified date order
            // RecursiveAll will get items in folders and sub-folders as well
            CamlQuery qry = new CamlQuery();
            qry.ViewXml = "<View Scope='RecursiveAll'>" +
                           "<ViewFields>" +
                           "<FieldRef Name='ID'/>" +
                           "<FieldRef Name='Title'/>" +
                           "<FieldRef Name='FileRef'/>" +
                           "<FieldRef Name='Created'/>" +
                           "<FieldRef Name='Author'/>" +
                           "<FieldRef Name='Modifed'/>" +
                           "<FieldRef Name='Editor'/>" +
                           "</ViewFields>" +
                           "<Query>" +
                           "<OrderBy>" +
                           "<FieldRef Name='Modified' Ascending='FALSE' />" +
                           "</OrderBy>" +
                           "</Query>" +
                           "</View>";


            // construt CAML query to retrieve items from Large Lists (i.e. lists that have more than 5000 list items)
            // we can't retrieve more than 5000 items as this will hit the List View Threshold. To get past this, we would have to increase the LVT setting in Central Admin
            // but increaasing LVT will cause performance issues. See this article for more details https://support.office.com/en-us/article/manage-large-lists-and-libraries-in-sharepoint-b8588dae-9387-48c2-9248-c24122f07c59 
            // for now we will retrieve 5000 items in descending Item Id order. this will get the items in modified date descending order (i.e. most recently modified items)
            CamlQuery largeListqry = new CamlQuery();
            largeListqry.ViewXml = "<View Scope='RecursiveAll'>" +
                          "<ViewFields>" +
                          "<FieldRef Name='ID'/>" +
                          "<FieldRef Name='Title'/>" +
                          "<FieldRef Name='FileRef'/>" +
                          "<FieldRef Name='Created'/>" +
                          "<FieldRef Name='Author'/>" +
                          "<FieldRef Name='Modified'/>" +
                          "<FieldRef Name='Editor'/>" +
                          "</ViewFields>" +
                          "<Query>" +
                          "<OrderBy Override='TRUE'>" +
                          "<FieldRef Name='ID' Ascending='FALSE' />" +
                          "</OrderBy>" +
                          "</Query>" +
                          "<RowLimit>5000</RowLimit>" +
                          "</View>";

            // get client context
            ClientContext context = GetOnPremiseClientContext(siteUrl, userName, userPassword, domain);

            Site site = context.Site;
            context.Load(site);
            context.ExecuteQuery();          
            

            // get collection of ALL Web URLs in the site collection
            // i.e. top level site and sub-sites
            IEnumerable<string> strAllWebUrls = site.GetAllWebUrls();
            foreach (string strUrl in strAllWebUrls)
            {
                Console.WriteLine("web url is {0}", strUrl);
                ClientContext webcontext = GetOnPremiseClientContext(strUrl, userName, userPassword, domain);
                Web web = webcontext.Web;
                webcontext.Load(web, w => w.Lists, w => w.Title);
                webcontext.ExecuteQuery();

                if (strUrl != siteUrl + "/Surveys")
                {
                    foreach (List list in web.Lists)
                    {
                        webcontext.Load(list, l => l.ItemCount, l => l.Title, l => l.DefaultViewUrl, l => l.DefaultDisplayFormUrl, l => l.BaseTemplate);
                        webcontext.ExecuteQuery();
                        // 101 => Documents ----   850 => Pages  ----- 100 => Generic List
                        Console.WriteLine("List info: {0} --- {1} --- {2} --- {3} ", list.Title, list.ItemCount, list.DefaultViewUrl, list.BaseTemplate);
                        int pos = Array.IndexOf(listExceptions, list.Title);
                        if (pos == -1 && list.ItemCount < 5000) //check if current list exists in exceptions list. If false, list will be inventoried
                        {
                            //Console.WriteLine(list.Title + " will be inventoried");
                            ListItemCollection collItems = list.GetItems(qry);
                            webcontext.Load(collItems);
                            webcontext.ExecuteQuery();
                            if (collItems.Count > 0)
                            {
                                int intGenericList = Array.IndexOf(genericListIds, list.BaseTemplate);

                                foreach (ListItem item in collItems)
                                {
                                    webcontext.Load(item);
                                    webcontext.ExecuteQuery();
                                    FieldUserValue authorValue = (FieldUserValue)item["Author"];
                                    FieldUserValue editorValue = (FieldUserValue)item["Editor"];
                                    string strAuthor = string.Empty;
                                    string strEditor = string.Empty;
                                    if (authorValue.LookupValue != null)
                                    {
                                        strAuthor = (authorValue.LookupValue).Replace(",", " ") + ",";
                                    }

                                    if (editorValue.LookupValue != null)
                                    {
                                        strEditor = (editorValue.LookupValue).Replace(",", " ") + ",";
                                    }

                                    string strItemTitle = string.Empty;
                                    if (item["Title"] != null)
                                    {
                                        strItemTitle = item["Title"].ToString().Replace(",", " ").Replace("\n", "");
                                    }
                                    string strItemUrl = string.Empty;
                                    if (intGenericList != -1)
                                    {
                                        strItemUrl = list.DefaultDisplayFormUrl + "?ID=" + item.Id.ToString();
                                    }
                                    else
                                    {
                                        strItemUrl = item["FileRef"].ToString();
                                    }
                                    string strListType = string.Empty;
                                    string strItemType = string.Empty;
                                    if (item.FileSystemObjectType.ToString() == "Folder")
                                    {
                                        string[] arrListAndItemType = GetListAndItemType(list.BaseTemplate, strItemUrl, "Folder");
                                        strListType = arrListAndItemType[0];
                                        strItemType = arrListAndItemType[1];
                                    }
                                    else //it is a File
                                    {
                                        string[] arrListAndItemType = GetListAndItemType(list.BaseTemplate, strItemUrl, "File");
                                        strListType = arrListAndItemType[0];
                                        strItemType = arrListAndItemType[1];

                                    }

                                    Console.WriteLine("Item info: {0} -- {1} -- {2} -- {3} -- {4} -- {5} -- {6}  -- {7} -- {8} -- {9} -- {10} ---{11} --- {12} --- {13}",
                                        web.Title, strUrl, list.Title, list.ItemCount.ToString(), list.DefaultViewUrl, strListType, item.Id.ToString(), strItemType, strItemTitle, item["Created"],
                                        strAuthor, item["Modified"], strEditor, strItemUrl);

                                    sbInvCSVFile.AppendLine(web.Title + "," +
                                                            strUrl + "," +
                                                            list.Title + "," +
                                                            list.ItemCount.ToString() + "," +
                                                            list.DefaultViewUrl + "," +
                                                            strListType + "," +
                                                            item.Id.ToString() + "," +
                                                            strItemType + "," +
                                                            strItemTitle + "," +
                                                            item["Created"] + "," +
                                                            strAuthor + "," +
                                                            item["Modified"] + "," +
                                                            strEditor + "," +
                                                            strItemUrl
                                                            );
                                }
                            }
                        }

                        if (pos == -1 && list.ItemCount > 5000) // This is a LARGE LIST (i.e. contains more than 5000 items. Log it for now. Will need to inventory these lists separately)
                        {
                            sbLargeListsCSVFile.AppendLine(web.Title + "," +
                                                           strUrl + "," +
                                                           list.Title + "," +
                                                           list.ItemCount.ToString() + "," +
                                                           list.DefaultViewUrl
                                                           );
                        }

                    }
                }

            }


            // process Large List now  -This list has more than 5000 items in it, and we can retrieve a MAX of 5000 items, without increasing List View Threshold setting
            ClientContext teamcontext = GetOnPremiseClientContext(siteUrl, userName, userPassword, domain);
            Web teamweb = teamcontext.Web;
            teamcontext.Load(teamweb, w => w.Lists, w => w.Title);
            teamcontext.ExecuteQuery();

            // get TeamSiteDirectory list and loop through each item in the list
            List listTeamSiteDir = teamweb.Lists.GetByTitle("TeamSiteDirectory");
            teamcontext.Load(listTeamSiteDir, l => l.Title, l => l.DefaultViewUrl, l => l.ItemCount, l => l.DefaultDisplayFormUrl, l => l.BaseTemplate);
            teamcontext.ExecuteQuery();

            ListItemCollection collTeamSiteItems = listTeamSiteDir.GetItems(largeListqry);
            teamcontext.Load(collTeamSiteItems);
            teamcontext.ExecuteQuery();
            if (collTeamSiteItems.Count > 0)
            {
                foreach (ListItem item in collTeamSiteItems)
                {
                    teamcontext.Load(item);
                    teamcontext.ExecuteQuery();

                    FieldUserValue authorValue = (FieldUserValue)item["Author"];
                    FieldUserValue editorValue = (FieldUserValue)item["Editor"];
                    string strItemTitle = string.Empty;
                    if (item["Title"] != null)
                    {
                        strItemTitle = item["Title"].ToString().Replace(",", " ").Replace("\n", "");
                    }

                    Console.WriteLine("Item info: {0} -- {1} -- {2} -- {3} -- {4} -- {5} -- {6}  -- {7} -- {8} -- {9} -- {10} ---- {11}",
                        teamweb.Title, siteUrl, listTeamSiteDir.Title, listTeamSiteDir.ItemCount.ToString(), listTeamSiteDir.DefaultViewUrl, item.Id.ToString(), strItemTitle, item["Created"],
                        authorValue.LookupValue, item["Modified"], editorValue.LookupValue, listTeamSiteDir.DefaultDisplayFormUrl + "?ID=" + item.Id.ToString());
                    sbInvCSVFile.AppendLine(teamweb.Title + "," +
                                            siteUrl + "," +
                                            listTeamSiteDir.Title + "," +
                                            listTeamSiteDir.ItemCount.ToString() + "," +
                                            listTeamSiteDir.DefaultViewUrl + "," +
                                            "List" + "," +
                                            item.Id.ToString() + "," +
                                            "Item" + "," +
                                            strItemTitle + "," +
                                            item["Created"] + "," +
                                            (authorValue.LookupValue).Replace(",", " ") + "," +
                                            item["Modified"] + "," +
                                            (editorValue.LookupValue).Replace(",", " ") + "," +
                                            listTeamSiteDir.DefaultDisplayFormUrl + "?ID=" + item.Id.ToString()
                                            );
                }
            }            


            //Write docs info to CSV file
            System.IO.File.AppendAllText(csvFilePath, sbInvCSVFile.ToString());  //this is the file that contains the site collection inventory
            System.IO.File.AppendAllText(csvLargeListsFilePath, sbLargeListsCSVFile.ToString()); //this is the file that contains info on Large Lists
            Console.WriteLine("===============================================================");
            Console.WriteLine("Site collection inventory DONE!");
            Console.WriteLine();

            watch.Stop();
            long elapsedMilliseconds = watch.ElapsedMilliseconds;
            double elapsedMinutes = TimeSpan.FromMilliseconds(elapsedMilliseconds).TotalMinutes;

            Console.WriteLine("Inventory report took {0} minutes to generate", elapsedMinutes);
            Console.ReadLine();

        }

        private static string[] GetListAndItemType(int baseTemplate, string strItemUrl, string strFolderOrFile)
        {
            string strListType = string.Empty;
            string strItemType = string.Empty;


            switch (baseTemplate)
            {
                case 100:
                    strListType = "List";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 101:
                    strListType = "Documents Library";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = GetFileExtension(strItemUrl);
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 102:
                    strListType = "Survey";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Survey Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }

                    break;
                case 103:
                    strListType = "Links";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 104:
                    strListType = "Announcement";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 105:
                    strListType = "Contact";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 106:
                    strListType = "Calendar";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 107:
                    strListType = "Task";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 108:
                    strListType = "Discussion Board";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 109:
                    strListType = "Picture Library";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = GetFileExtension(strItemUrl);
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 850:
                    strListType = "Pages Library";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Page";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                case 851:
                    strListType = "Images";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = GetFileExtension(strItemUrl);
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
                default:
                    strListType = "List";
                    if (strFolderOrFile == "File")
                    {
                        strItemType = "Item";
                    }
                    else
                    {
                        strItemType = "Folder";
                    }
                    break;
            }
            string[] strArrListTypeAndItemType = new string[] { strListType, strItemType };
            return strArrListTypeAndItemType;
        }

        private static string GetFileExtension(string strItemUrl)
        {
            string strItemType = string.Empty;
            string strExtension = System.IO.Path.GetExtension(strItemUrl).ToLowerInvariant().Replace(".", "");
            
            switch (strExtension)
            {
                case "pdf":
                    strItemType = "PDF";
                    break;
                case "doc":
                case "docx":
                case "docm":
                    strItemType = "Document";
                    break;
                case "ppt":
                case "pptx":
                case "pptm":
                    strItemType = "Presentation";
                    break;
                case "xls":
                case "xlsx":
                    strItemType = "Spreadsheet";
                    break;
                case "jpg":
                case "jpeg":
                case "gif":
                case "png":
                case "tif":
                case "tiff":
                case "pcd":
                    strItemType = "Image";
                    break;
                case "ics":
                    strItemType = "Calendar";
                    break;
                case "swf":
                    strItemType = "Shockwave Flash";
                    break;
                case "avi":
                case "flv":
                case "wmv":
                case "mp4":
                case "mpeg":
                case "mov":
                    strItemType = "Video";
                    break;
                default:
                    strItemType = "Item";
                    break;
            }
            return strItemType;
        }


        // use this method when running this in a SharePoint on-premise environment, to get the ClientContext object
        public static ClientContext GetOnPremiseClientContext(string siteCollUrl, string username, string password, string domain)
        {
            //initialization
            ClientContext context = new OfficeDevPnP.Core.AuthenticationManager().GetNetworkCredentialAuthenticatedContext(siteCollUrl, username, password, domain);

#if !DEBUG
                    //hook request exector
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        //https://ravisoftltd.wordpress.com/2013/09/12/connect-to-sharepoint-siteconfigured-with-claim-based-authentication-with-managed-csom/ 
                        e.WebRequestExecutor.RequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                    };
#endif

            //return
            return context;
        }
    }
}
