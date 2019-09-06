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
            List<Folder> allFolders = new List<Folder>();

            //Build the allFolders list
            foreach (Folder f in childFolders)
            {
                if ((String.Equals(f.DisplayName, "Inbox")) || (String.Equals(f.DisplayName, "Conversation History"))
                    || (String.Equals(f.DisplayName, "Save")) || (String.Equals(f.DisplayName, "Trips")) || (String.Equals(f.DisplayName, "Sent Items")))
                {
                    allFolders.Add(f);
                    if (f.ChildFolderCount > 0)
                    {
                        FolderView folderView = new FolderView(f.ChildFolderCount);
                        GetSubFolders(f, folderView, allFolders);
                    }
                }
            }

            //Get (and index) all messages in all folders
            foreach (Folder f in allFolders)
            {
                GetMessages(f);
                //Console.WriteLine(f.DisplayName);
                //Console.WriteLine(f.TotalCount);
                //Console.WriteLine();
            }
            Console.ReadLine();
        }

        public static void GetSubFolders(Folder folder, FolderView folderView, List<Folder> allFolders)
        {
            FindFoldersResults subFolders = folder.FindFolders(folderView);
            allFolders.AddRange(subFolders.Folders);

            foreach (Folder f in subFolders)
            {
                if (f.ChildFolderCount > 0)
                {
                    FolderView folderViewSub = new FolderView(f.ChildFolderCount);
                    GetSubFolders(f, folderViewSub, allFolders);
                }
            }
        }

        public static void GetMessages(Folder folder)
        {
            int itemCount = folder.TotalCount;
            FindItemsResults<Item> results = null;

            //for (int i = 0; i < itemCount; i += 200)
            for (int i = 4400; i < 4600; i += 200)
            {
                Console.WriteLine(i.ToString());
                List<MessageIndexObject> messageIndexObjects = new List<MessageIndexObject>();
                ItemView view = new ItemView(200, i);
                results = service.FindItems(folder.Id, view);
                
                //if (results.Items.Count > 0)
                if (results.Items.Count > 0)
                {
                    PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.TextBody);
                    ExtendedPropertyDefinition extendedProps = new ExtendedPropertyDefinition(0x0FFF, MapiPropertyType.Binary);
                    propSet.Add(ItemSchema.StoreEntryId);
                    propSet.Add(extendedProps);
                    service.LoadPropertiesForItems(results, propSet);
                    foreach (Item k in results)
                    {
                        Byte[] EntryVal = null;
                        String HexEntryId = "";
                        if (k.TryGetProperty(extendedProps, out EntryVal))
                        {
                            HexEntryId = BitConverter.ToString(EntryVal).Replace("-", "");
                        }
                        messageIndexObjects.Add(new MessageIndexObject(k.Subject, k.TextBody, HexEntryId));
                    }
                    //Push 200 messages for indexing at a time
                    //Indexer indexer = new Indexer();
                    //indexer.PushMessages(messageIndexObjects);
                }
            }
        }
    }
}
