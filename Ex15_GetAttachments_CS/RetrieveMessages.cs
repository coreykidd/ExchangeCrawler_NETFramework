using System;
using System.IO;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeCrawl
{
    class RetrieveMessages
    {
        static IUserData UserData = UserDataFromConsole.GetUserData();
        static ExchangeService service = Service.ConnectToService(UserData, new TraceListener());

        static void Main(string[] args)
        {
            Folder rootFolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);
            FindFoldersResults childFolders = rootFolder.FindFolders(new FolderView(rootFolder.ChildFolderCount));
            foreach (Folder f in childFolders)
            {
                if (String.Equals(f.DisplayName, "Inbox"))
                {
                    Folder inbox = f;
                    //FindItemsResults<Item> items = GetMessages(inbox);
                    GetMessages(inbox);
                    //foreach (Item i in items)
                    //{
                    //    Console.WriteLine(i.Subject);
                    //}
                    //Console.ReadLine();
                }
            }
            Console.ReadLine();
        }

        //public static FindItemsResults<Item> GetMessages(Folder folder)
        public static void GetMessages(Folder folder)
        {
            int itemCount = folder.TotalCount;
            FindItemsResults<Item> results = null;

            int j = 1;
            for (int i = 0; i < 30; i+=10)
            {
                ItemView view = new ItemView(10, i);
                results = service.FindItems(folder.Id, view);
                PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.TextBody);
                service.LoadPropertiesForItems(results, propSet);

                foreach (Item k in results)
                {
                    Console.WriteLine($"{j.ToString()}. {k.Subject}");
                    Console.WriteLine(k.TextBody);
                    Console.WriteLine(k.ConversationId);
                    j++;
                }
            }
        }
    }
}
