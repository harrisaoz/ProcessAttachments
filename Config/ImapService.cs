using System.Configuration;

namespace Config
{
    public class ImapService : ConfigurationSection
    {
        [ConfigurationProperty("hostname", IsRequired = true)]
        public string Hostname
        {
            get { return (string)this["hostname"]; }
            set { this["hostname"] = value; }
        }

        [ConfigurationProperty("port", IsRequired = false, DefaultValue = 993)]
        public int Port
        {
            get { return (int)this["port"]; }
            set { this["port"] = value; }
        }
    }
}
