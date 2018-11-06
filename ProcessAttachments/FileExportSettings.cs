using System.Configuration;

namespace ProcessAttachments
{
    internal class FileExportSettings : ConfigurationSection
    {
        [ConfigurationProperty("exportDir", IsRequired = true)]
        public string ExportDirectory
        {
            get { return (string)this["exportDir"]; }
            set { this["exportDir"] = value; }
        }
    }
}
