using System.Configuration;

namespace Config
{
    public class ImapFolders : ConfigurationSection
    {
        [ConfigurationProperty("sourceFolder", IsRequired = true)]
        public string SourceFolder
        {
            get { return (string)this["sourceFolder"]; }
            set { this["sourceFolder"] = value; }
        }

        [ConfigurationProperty("processedFolder", IsRequired = false)]
        public string ProcessedFolder
        {
            get { return (string)this["processedFolder"]; }
            set { this["processedFolder"] = value; }
        }

        [ConfigurationProperty("attentionFolder", IsRequired = false)]
        public string AttentionFolder
        {
            get { return (string)this["attentionFolder"]; }
            set { this["attentionFolder"] = value; }
        }
    }
}
