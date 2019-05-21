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

        [ConfigurationProperty("port", IsRequired = false, DefaultValue = (uint)993)]
        public uint Port
        {
            get { return (uint)this["port"]; }
            set { this["port"] = value; }
        }
    }
}
