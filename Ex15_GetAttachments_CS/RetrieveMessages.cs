using System;
using System.IO;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;

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
            List<MessageIndexObject> messageIndexObjects = new List<MessageIndexObject>();
            foreach (Folder f in childFolders)
            {
                if (String.Equals(f.DisplayName, "Inbox"))
                {
                    Folder inbox = f;
                    messageIndexObjects = GetMessages(inbox);
                }
            }
            Indexer indexer = new Indexer();
            indexer.PushMessages(messageIndexObjects);
        }

        public static List<MessageIndexObject> GetMessages(Folder folder)
        {
            int itemCount = folder.TotalCount;
            FindItemsResults<Item> results = null;
            List<MessageIndexObject> messageIndexObjects = new List<MessageIndexObject>();

            //int j = 1;
            for (int i = 0; i < 100; i+=10)
            {
                ItemView view = new ItemView(10, i);
                results = service.FindItems(folder.Id, view);
                PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.TextBody);
                service.LoadPropertiesForItems(results, propSet);

                foreach (Item k in results)
                {
                    //Console.WriteLine($"{j.ToString()}. {k.Subject}");
                    //Console.WriteLine(k.TextBody);
                    //Console.WriteLine(k.ConversationId);
                    //j++;
                    messageIndexObjects.Add(new MessageIndexObject(k.Subject, k.TextBody, k.ConversationId));
                }
            }
            return messageIndexObjects;
        }
    }
}
