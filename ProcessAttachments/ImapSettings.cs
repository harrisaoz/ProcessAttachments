using System.Configuration;

namespace ProcessAttachments
{
    class ImapSettings : ConfigurationSection
    {
        [ConfigurationProperty("hostname", IsRequired = true)]
        public string Hostname
        {
            get { return (string)this["hostname"]; }
            set { this["hostname"] = value; }
        }

        [ConfigurationProperty("username", IsRequired = true)]
        public string Username
        {
            get { return (string)this["username"]; }
            set { this["username"] = value; }
        }

        [ConfigurationProperty("password", IsRequired = true)]
        public string Password
        {
            get { return (string)this["password"];  }
            set { this["password"] = value; }
        }

        [ConfigurationProperty("sourceFolder", IsRequired = true)]
        public string SourceFolder
        {
            get { return (string)this["sourceFolder"]; }
            set { this["sourceFolder"] = value; }
        }

        [ConfigurationProperty("processedFolder", IsRequired = true)]
        public string ProcessedFolder
        {
            get { return (string)this["processedFolder"]; }
            set { this["processedFolder"] = value; }
        }

        [ConfigurationProperty("attentionFolder", IsRequired = true)]
        public string AttentionFolder
        {
            get { return (string)this["attentionFolder"]; }
            set { this["attentionFolder"] = value; }
        }
    }
}
