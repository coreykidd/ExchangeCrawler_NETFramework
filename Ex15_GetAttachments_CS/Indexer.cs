using Nest;
using System;
using System.Collections.Generic;

namespace ExchangeCrawl
{
    public class Indexer
    {
        public void PushMessages(List<MessageIndexObject> messageIndexObjects)
        {
            var settings = new ConnectionSettings(new Uri("http://localhost:9200")).DefaultIndex("emails");
            var client = new ElasticClient(settings);

            foreach(MessageIndexObject messageIndexObject in messageIndexObjects)
            {
                client.IndexDocument(messageIndexObject);
            }
        }
    }
}
