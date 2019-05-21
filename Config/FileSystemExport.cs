using System.Configuration;

namespace Config
{
    public class FileSystemExport : ConfigurationSection
    {
        [ConfigurationProperty("exportDir", IsRequired = true)]
        public string ExportDirectory
        {
            get { return (string)this["exportDir"]; }
            set { this["exportDir"] = value; }
        }
    }
}
