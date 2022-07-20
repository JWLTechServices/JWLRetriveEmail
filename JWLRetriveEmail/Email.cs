using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JWLRetriveEmail
{
    [Serializable]
    public class Email
    {
        public Email()
        {
            this.Attachments = new List<Attachments>();
        }
        public int MessageNumber { get; set; }
        public string From { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime DateSent { get; set; }
        public List<Attachments> Attachments { get; set; }
    }

    [Serializable]
    public class Attachments
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public byte[] Content { get; set; }
    }
}
