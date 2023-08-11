using System.Collections.Generic;

namespace ExportViewUtility.Models
{
    public class Configuration
    {
        public string crmConnectionString { get; set; }
        public string fetchXml { get; set; }
        public string fetchXmlFile { get; set; }
        public List<ColumnMapping> columnMapping { get; set; }
        public List<string> emailAddress { get; set; }
        public MailConfiguration mailServer { get; set; }
        public string subject { get; set; }
        public string fileName { get; set; }
        public string body { get; set; }
        public bool skipIfEmpty { get; set; }
    }
}
