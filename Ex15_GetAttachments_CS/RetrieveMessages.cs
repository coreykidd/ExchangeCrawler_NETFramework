using System;
using System.IO;
using System.Collections.ObjectModel;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeCrawl
{
    class RetrieveMessages
    {
        static IUserData UserData = UserDataFromConsole.GetUserData();
        static ExchangeService service = Service.ConnectToService(UserData, new TraceListener());

        static void Main(string[] args)
        {
            //FindItemsResults<Item> items = GetMessages();
            //foreach(Item i in items)
            //{
            //    Console.WriteLine(i.Subject);
            //}
            //Console.ReadLine();

            Folder folder = GetFolder();
            FindFoldersResults childFolders = folder.FindFolders(new FolderView(folder.ChildFolderCount));
        }

        public static FindItemsResults<Item> GetMessages()
        {
            ItemView view = new ItemView(10);
            FindItemsResults<Item> results = service.FindItems(WellKnownFolderName.Inbox, view);
            PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.TextBody);
            service.LoadPropertiesForItems(results, propSet);
            return results;
        }

        public static Folder GetFolder()
        {
            Folder rootfolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);
            return rootfolder;
        }
    }
}
