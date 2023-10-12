using Microsoft.SharePoint.Client;
using System;
using System.Configuration;

namespace CSOM_Copy_Files
{
    class Program
    {
        static void Main(string[] args)
        {
            GetDocumentInventory();
            //CopyDocuments("https://fightercorks.sharepoint.com/", "https://fightercorks.sharepoint.com/sites/hr", "sourcelibrary", "destinationlib");
        }
        static void CopyDocuments(string srcUrl, string destUrl, string srcLibrary, string destLibrary)
        {
            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext clientContext = authManager.GetWebLoginClientContext(srcUrl);
            ClientContext srcContext = new ClientContext(srcUrl);

            ClientContext destContext = new ClientContext(destUrl);

            Web srcWeb = srcContext.Web;

            List srcList = srcWeb.Lists.GetByTitle(srcLibrary);

            Web destWeb = destContext.Web;

            destContext.Load(destWeb);

            destContext.ExecuteQuery();
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
    }
}
