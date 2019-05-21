using System.Configuration;

namespace Config
{
    public class ConfigLoader
    {
        private Configuration config;

        public ConfigLoader()
        {
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        }

        public ImapFolders imapFolderSettings()
        {
            return config.GetSection("imapFolders") as ImapFolders;
        }

        public ImapService imapServiceSettings()
        {
            return config.GetSection("imapService") as ImapService;
        }

        public FileSystemExport fileSystemExportSettings()
        {
            return config.GetSection("fileSystemExport") as FileSystemExport;
        }

        public Credentials credentialsSettings()
        {
            return config.GetSection("credentials") as Credentials;
        }
    }
}
