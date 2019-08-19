using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeCrawl
{
    public class MessageIndexObject
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string ConversationId { get; set; }

        public MessageIndexObject(string subject, string body, string conversationId)
        {
            this.Subject = subject;
            this.Body = body;
            this.ConversationId = conversationId;
        }
    }
}
