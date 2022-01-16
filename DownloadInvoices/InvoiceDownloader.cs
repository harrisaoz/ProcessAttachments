using System;
using System.Linq;

using MailKit;

using Config;
using ImapAttachmentProcessing;

namespace DownloadInvoices
{
    class InvoiceDownloader
    {
        static void Main(string[] args)
        {
            var configLoader = new ConfigLoader();

            var downloader = new InvoiceDownloader(
                configLoader.imapServiceSettings(),
                configLoader.credentialsSettings(),
                configLoader.imapFolderSettings().SourceFolder,
                new AttachmentProcessorBuilder()
                    .PermitAttachmentType(CommonContentTypes.PDF_NORMAL)
                    .PermitAttachmentType(CommonContentTypes.PDF_GENERAL)
                    .IgnoringAttachments((contentType, fileName) => fileName
                        .Contains("detalhe") || fileName
                        .EndsWith(".zip")
                    )
                    .NameDownloadsByFolder(FileNamers.FolderBasedName)
                    .DownloadTo(configLoader.fileSystemExportSettings().ExportDirectory)
                    .IncludeMessageHeaders(
                            MessageSummaryItems.UniqueId |
                            MessageSummaryItems.Envelope |
                            MessageSummaryItems.GMailLabels |
                            MessageSummaryItems.BodyStructure |
                            MessageSummaryItems.Size |
                            MessageSummaryItems.ModSeq
                    )
                    .Build()
            );

            downloader.TryRun(downloader.RunImplementation);
        }

        readonly SimpleImapService imapService;
        private readonly string sourceFolderName;
        private readonly AttachmentProcessor attachmentProcessor;
        private readonly MailFolderFinder mailFolderFinder;

        private InvoiceDownloader(
            ImapService serviceConfig,
            Credentials credentials,
            string folder,
            AttachmentProcessor attachmentProcessor)
        {
            imapService = new SimpleImapService(serviceConfig, credentials);
            sourceFolderName = folder;
            mailFolderFinder = new MailFolderFinder(
                client => folderName => client
                    .GetFolder(client.PersonalNamespaces[0])
                    .GetSubfolders(false)
                    .Where(subFolder => subFolder.Name.ToUpper().Equals(folderName.ToUpper())),
                imapService.client
            );

            this.attachmentProcessor = attachmentProcessor;
        }

        public void TryRun(Action run) {
            try
            {
                imapService.Connect();

                run();
            }
            catch (Exception e)
            {
                Console.WriteLine("Extraction of attachments failed. {0}", e);
                return;
            }
            finally
            {
                imapService.Disconnect();
            }

            Console.WriteLine("Program terminating normally.");
        }

        public void RunImplementation()
        {
            int emailsProcessedCount = mailFolderFinder
                .FindByName(this.sourceFolderName)
                .Select(folder => TraverseFolder(folder))
                .Sum();
        }

        int TraverseFolder(IMailFolder mailFolder)
        {
            return attachmentProcessor.TryVisitFolder(mailFolder, (processed, requireAttention) => {}) +
                mailFolder
                    .GetSubfolders(subscribedOnly: false)
                    .Select(subFolder => TraverseFolder(subFolder))
                    .Sum();
        }
    }
}
