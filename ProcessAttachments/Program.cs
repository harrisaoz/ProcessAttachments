using System;
using System.Configuration;

namespace ProcessAttachments
{
    class Program
    {
        static void Main(string[] args)
        {
            //config();
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var imapSettings = config.GetSection("imapSettings") as ImapSettings;
            var fileExportSettings = config.GetSection("fileExportSettings") as FileExportSettings;

            if (args[0].Equals("invoices"))
            {
                InvoiceDownloader.execute(imapSettings, fileExportSettings);
            }
            else
            {
                TimesheetProcessor.execute(imapSettings, fileExportSettings);
            }
        }

        static void config()
        {
            var appSettings = ConfigurationManager.AppSettings;

            foreach (var key in appSettings.AllKeys)
            {
                Console.WriteLine("(key = {0}) (value = {1})", key, appSettings[key]);
            }

            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            Console.WriteLine("Config file: {0}", config.FilePath);

            var imapSettings = config.GetSection("imapSettings") as ImapSettings;
            var fileExportSettings = config.GetSection("fileExportSettings") as FileExportSettings;

            Console.WriteLine("{0} = {1}", "hostname", imapSettings.Hostname);
            Console.WriteLine("{0} = {1}", "username", imapSettings.Username);
            Console.WriteLine("{0} = {1}", "password", imapSettings.Password);

            Console.WriteLine("{0} = {1}", "exportDir", fileExportSettings.ExportDirectory);

            config.Save(ConfigurationSaveMode.Modified);
        }
    }
}
