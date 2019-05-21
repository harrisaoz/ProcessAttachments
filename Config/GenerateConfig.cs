using System.Configuration;

namespace Config
{
    class GenerateConfig
    {
        static void Main(string[] args)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var imapServiceSection = new ImapService();
            imapServiceSection.Hostname = "default";
            imapServiceSection.Port = 993;

            config.Sections.Add("imapService", imapServiceSection);

            var credentialsSection = new Credentials();
            credentialsSection.Username = "username";
            credentialsSection.Password = "";
            config.Sections.Add("credentials", credentialsSection);

            var filesystemExportSection = new FileSystemExport();
            filesystemExportSection.ExportDirectory = "";
            config.Sections.Add("fileSystemExport", filesystemExportSection);

            var imapFoldersSection = new ImapFolders();
            imapFoldersSection.SourceFolder = "/";
            imapFoldersSection.ProcessedFolder = "PROCESSED";
            imapFoldersSection.AttentionFolder = "ATTENTION";
            config.Sections.Add("imapFolders", imapFoldersSection);

            config.Save(ConfigurationSaveMode.Full);
        }
    }
}
