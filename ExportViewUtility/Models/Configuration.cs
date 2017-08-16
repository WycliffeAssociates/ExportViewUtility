using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportViewUtility.Models
{
    public class Configuration
    {
        public string crmConnectionString;
        public string fetchXml;
        public string fetchXmlFile;
        public List<ColumnMapping> columnMapping;
        public List<string> emailAddress;
        public MailConfiguration mailServer;
        public string subject;
        public string fileName;
        public string body;
        public bool skipIfEmpty;
    }
}
