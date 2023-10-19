using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Linq;

namespace CSOM_Copy_Files
{
    class Program
    {
        static void Main(string[] args)
        {
            importExcelDataIntoList();
            //copyListItems();
            //GetAllListItemsWithPagination()
            //GetDocumentInventory();

        }
        static void copyListItems()
        {
            string siteUrl = Convert.ToString(ConfigurationManager.AppSettings["sourceSite"]);
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext clientContext = authManager.GetWebLoginClientContext(siteUrl);
            ClientContext cc = authManager.GetWebLoginClientContext(siteUrl);
            List list = clientContext.Web.Lists.GetByTitle("source");
            List listDest = clientContext.Web.Lists.GetByTitle("destination");
            ListItemCollection collListItem = null;

            try
            {
                clientContext.Load(list);
                clientContext.Load(listDest);

                clientContext.ExecuteQuery();

                CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();
                collListItem = list.GetItems(camlQuery);
                clientContext.Load(collListItem);
                clientContext.ExecuteQuery();

                if (collListItem != null && collListItem.Count > 0)
                {
                    foreach (ListItem oListItem in collListItem)
                    {
                        ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                        ListItem listItemCreation = listDest.AddItem(itemInfo);
                        listItemCreation["Title"] = Convert.ToString(oListItem["Title"]);
                        listItemCreation["sourceid"] = oListItem["ID"];
                        listItemCreation["credit"] = oListItem["credit"];
                        listItemCreation["debit"] = oListItem["debit"];
                        listItemCreation.Update();
                        clientContext.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        static void CopyDocuments(string srcUrl, string destUrl, string srcLibrary, string destLibrary)
        {
            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
           // ClientContext clientContext = authManager.GetWebLoginClientContext(srcUrl);
            ClientContext srcContext = new ClientContext(srcUrl);

            ClientContext destContext = new ClientContext(destUrl);

            Web srcWeb = srcContext.Web;
            List srcList = srcWeb.Lists.GetByTitle(srcLibrary);
            Web destWeb = destContext.Web;
            destContext.Load(destWeb);
            destContext.ExecuteQuery();
            srcContext.ExecuteQuery();
            try
            {
                Microsoft.SharePoint.Client.File file = srcContext.Web.GetFileByServerRelativeUrl("/sites/Publishing/test.pdf");
                srcContext.Load(file);
                srcContext.ExecuteQuery();
                string location = destWeb.ServerRelativeUrl.TrimEnd('/') + "/" + destLibrary.Replace(" ", "") + "/" + file.Name;
                FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(srcContext, file.ServerRelativeUrl);
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(destContext, location, fileInfo.Stream, true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        static void GetDocumentInventory()
        {
            try
            {
                string siteUrl = Convert.ToString(ConfigurationManager.AppSettings["sourceSite"]);
                var authManager = new OfficeDevPnP.Core.AuthenticationManager();
                ClientContext context = authManager.GetWebLoginClientContext(siteUrl);

                var list = context.Web.Lists.GetByTitle("Folder Templates");
                var rootFolder = list.RootFolder;
                GetFoldersAndFiles(rootFolder, context);


                //FolderCollection folders = list.RootFolder.Folders;
                //context.Load(folders);
                //context.ExecuteQuery();
                //foreach (var folder in folders)
                //{
                //    GetFoldersAndFiles(folder, context);
                //}
                Console.ReadLine();
            }
            catch (Exception ex)
            { throw ex; }
        }
        private static void GetFoldersAndFiles(Folder mainFolder, ClientContext clientContext)
        {
            try
            {
                clientContext.Load(mainFolder, k => k.Name, k => k.Files, k => k.Folders, k => k.ServerRelativeUrl);
                clientContext.ExecuteQuery();
                foreach (var folder in mainFolder.Folders)
                {
                    Console.WriteLine(folder.ServerRelativeUrl);
                    GetFoldersAndFiles(folder, clientContext);
                }
                foreach (var file in mainFolder.Files)
                {
                    Console.WriteLine(file.ServerRelativeUrl);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static List<ListItem> GetAllListItemsWithPagination()
        {
            List<ListItem> items = new List<ListItem>();
            string siteUrl = Convert.ToString(ConfigurationManager.AppSettings["sourceSite"]);
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext clientContext = authManager.GetWebLoginClientContext(siteUrl);
            List list = clientContext.Web.Lists.GetByTitle("source");
            try
            {
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                ListItemCollectionPosition position = null;
                int rowLimit = 50;
                var camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
                <Query>
                    <OrderBy Override='TRUE'><FieldRef Name='ID'/></OrderBy>
                </Query>
                <ViewFields>
                    <FieldRef Name='Title'/><FieldRef Name='Modified' /><FieldRef Name='Editor' />
                </ViewFields>
                <RowLimit Paged='TRUE'>" + rowLimit + "</RowLimit></View>";
                do
                {
                    ListItemCollection listItems = null;
                    camlQuery.ListItemCollectionPosition = position;
                    listItems = list.GetItems(camlQuery);
                    clientContext.Load(listItems);
                    clientContext.ExecuteQuery();
                    position = listItems.ListItemCollectionPosition;
                    items.AddRange(listItems.ToList());
                }
                while (position != null);
            }
            catch (Exception ex)
            {
            }
            return items;
        }
        static void importExcelDataIntoList()
        {
            try
            {
                string conn =
                  @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\empData.xlsx;" +
                  @"Extended Properties='Excel 12.0;HDR=Yes;'";

                string joiningDate;

                string siteUrl = Convert.ToString(ConfigurationManager.AppSettings["sourceSite"]);
                var authManager = new OfficeDevPnP.Core.AuthenticationManager();
                ClientContext clientContext = authManager.GetWebLoginClientContext(siteUrl);
                ClientContext cc = authManager.GetWebLoginClientContext(siteUrl);
                ClientContext ccManager = authManager.GetWebLoginClientContext(siteUrl);
                ClientContext ccEmployee = authManager.GetWebLoginClientContext(siteUrl);
                List list = clientContext.Web.Lists.GetByTitle("Employee Directory");
                User manager,employee = null;
                ListItemCollection collListItem = null;
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                using (OleDbConnection connection = new OleDbConnection(conn))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                    {
                        int i = 1;
                        while (dr.Read())
                        {
                            collListItem = null;
                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                                "<Value Type='Text'>" + Convert.ToString(dr["Employee ID"]) + "</Value></Eq></Where></Query></View>";
                            collListItem = list.GetItems(camlQuery);
                            clientContext.Load(collListItem);
                            clientContext.ExecuteQuery();
                            if (collListItem != null && collListItem.Count > 0)
                            {
                                foreach (ListItem oListItem in collListItem)
                                {
                                    manager = null;
                                    oListItem["employeeid"] = Convert.ToString(dr["Employee ID"]);
                                    oListItem["status"] = Convert.ToString(dr["Status"]);
                                    oListItem["firstname"] = Convert.ToString(dr["First Name"]);
                                    oListItem["lastname"] = Convert.ToString(dr["Last Name"]);
                                    oListItem["department"] = Convert.ToString(dr["Department"]);

                                    joiningDate = Convert.ToString(dr["Joining Date"]);
                                    if (!string.IsNullOrEmpty(joiningDate))
                                    {
                                        oListItem["joiningdate"] = joiningDate;
                                    }
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Manager Email"])))
                                    {
                                        try
                                        {
                                            manager = ccManager.Web.EnsureUser(Convert.ToString(dr["Manager Email"]));
                                            ccManager.Load(manager);
                                            ccManager.ExecuteQuery();
                                            oListItem["manager"] = manager.Id;
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }
                                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Employee Email"])))
                                    {
                                        try
                                        {
                                            employee = ccEmployee.Web.EnsureUser(Convert.ToString(dr["Employee Email"]));
                                            ccEmployee.Load(employee);
                                            ccEmployee.ExecuteQuery();
                                            oListItem["employee"] = employee.Id;
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }
                                    oListItem.Update();
                                    clientContext.ExecuteQuery();

                                    //oListItem.BreakRoleInheritance(true, false);
                                    //clientContext.ExecuteQuery();
                                    //try
                                    //{
                                    //    var user_group = clientContext.Web.SiteGroups.GetByName("HRGROUP");
                                    //    oListItem.RoleAssignments.GetByPrincipal(user_group).DeleteObject();
                                    //    clientContext.ExecuteQuery();
                                    //}
                                    //catch (Exception ex)
                                    //{
                                    //}
                                    //string hrbpgroup = Convert.ToString(dr["HRBP"]);
                                    //if (!string.IsNullOrEmpty(hrbpgroup))
                                    //{
                                    //    var hrbp_Group = clientContext.Web.SiteGroups.GetByName(hrbpgroup);
                                    //    var roleDefCol = new RoleDefinitionBindingCollection(clientContext);
                                    //    roleDefCol.Add(clientContext.Web.RoleDefinitions.GetByType(RoleType.Contributor));
                                    //    oListItem.RoleAssignments.Add(hrbp_Group, roleDefCol);
                                    //    oListItem["grp_HRBP"] = hrbp_Group;
                                    //    oListItem.Update();
                                    //    clientContext.ExecuteQuery();
                                    //}
                                }
                            }
                            else
                            {
                                ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                                ListItem listItemCreation = list.AddItem(itemInfo);
                                listItemCreation["employeeid"] = Convert.ToString(dr["Employee ID"]);
                                listItemCreation["status"] = Convert.ToString(dr["Status"]);
                                listItemCreation["firstname"] = Convert.ToString(dr["First Name"]);
                                listItemCreation["lastname"] = Convert.ToString(dr["Last Name"]);
                                listItemCreation["department"] = Convert.ToString(dr["Department"]);

                                joiningDate = Convert.ToString(dr["Joining Date"]);
                                if (!string.IsNullOrEmpty(joiningDate))
                                {
                                    listItemCreation["joiningdate"] = joiningDate;
                                }
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Manager Email"])))
                                {
                                    try
                                    {
                                        manager = ccManager.Web.EnsureUser(Convert.ToString(dr["Manager Email"]));
                                        ccManager.Load(manager);
                                        ccManager.ExecuteQuery();
                                        listItemCreation["manager"] = manager.Id;
                                    }
                                    catch (Exception ex)
                                    {
                                    }
                                }
                                if (!string.IsNullOrEmpty(Convert.ToString(dr["Employee Email"])))
                                {
                                    try
                                    {
                                        employee = ccEmployee.Web.EnsureUser(Convert.ToString(dr["Employee Email"]));
                                        ccEmployee.Load(employee);
                                        ccEmployee.ExecuteQuery();
                                        listItemCreation["employee"] = employee.Id;
                                    }
                                    catch (Exception ex)
                                    {
                                    }
                                }
                                listItemCreation.Update();
                                clientContext.ExecuteQuery();                               
                            }
                            Console.Write(i.ToString() + " " + Convert.ToString(dr["Employee Email"]));
                            Console.WriteLine();
                            i++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
            }
            Console.ReadLine();
        }
    }
}
