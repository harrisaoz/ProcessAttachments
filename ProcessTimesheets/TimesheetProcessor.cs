using System;
using System.Collections.Generic;

using MailKit.Search;
using MailKit;

using Config;
using ImapAttachmentProcessing;

namespace ProcessTimesheets
{
    class TimesheetProcessor
    {
        static void Main(string[] args)
        {
            var configLoader = new ConfigLoader();

            var processor = new TimesheetProcessor(
                configLoader.imapServiceSettings(),
                configLoader.credentialsSettings(),
                configLoader.imapFolderSettings().SourceFolder,
                configLoader.imapFolderSettings().ProcessedFolder,
                configLoader.imapFolderSettings().AttentionFolder,
                new AttachmentProcessorBuilder()
                    .PermitAttachmentTypes(CommonContentTypes.IMAGES)
                    .PermitAttachmentTypes(CommonContentTypes.PDF)
                    .PermitAttachmentType(CommonContentTypes.WORD_DOC)
                    .IgnoringAttachments((contentType, fileName) => contentType != null && (
                        contentType.IsMimeType("text", "plain") && (fileName == null || fileName.EndsWith(".txt")) ||
                        contentType.IsMimeType("text", "html") && (fileName == null || fileName.Contains(".htm")))
                    )
                    .DownloadTo(configLoader.fileSystemExportSettings().ExportDirectory)
                    .NameDownloadsBySummaryOnly(FileNamers.SummaryBasedName)
                    .UsingSearchQuery(
                        SearchQuery.DeliveredAfter(DateTime.Parse("2018-10-01"))
                    )
                    .IncludeMessageHeaders(
                        MessageSummaryItems.UniqueId |
                        MessageSummaryItems.Envelope |
                        MessageSummaryItems.BodyStructure |
                        MessageSummaryItems.InternalDate
                    )
                    .Build()
            );

            processor.TryRun();
        }

        private readonly SimpleImapService imapService;
        private readonly string sourceFolderName;
        private readonly string processedFolderName;
        private readonly string attentionFolderName;
        private readonly AttachmentProcessor attachmentProcessor;
        private readonly IMailFolderFinder mailFolderFinder;

        TimesheetProcessor(
            ImapService serviceConfig,
            Credentials credentials,
            string sourceFolder, string processedFolder, string attentionFolder,
            AttachmentProcessor attachmentProcessor)
        {
            imapService = new SimpleImapService(serviceConfig, credentials);
            sourceFolderName = sourceFolder;
            processedFolderName = processedFolder;
            attentionFolderName = attentionFolder;
            mailFolderFinder = new MailFolderFinder(
                client => folderName =>
                {
                    var folders = new List<IMailFolder>();
                    folders.Add(client.GetFolder(folderName));
                    return folders;
                },
                imapService.client
            );

            this.attachmentProcessor = attachmentProcessor;
        }

        Func<IMailFolder, IMailFolder, IMailFolder, Action<List<UniqueId>, List<UniqueId>>> onCompletion =
            (sourceFolder, processedFolder, attentionFolder) =>
            (List<UniqueId> processed, List<UniqueId> requireAttention) =>
            {
                processed.ForEach(
                    uid => Console.WriteLine("{0} processed", uid)
                );

                requireAttention.ForEach(
                    uid => Console.WriteLine("{0} requires attention", uid));

                sourceFolder.Close();
                sourceFolder.Open(FolderAccess.ReadWrite);
                sourceFolder.MoveTo(processed, processedFolder);
                sourceFolder.MoveTo(requireAttention, attentionFolder);
            };

        public void TryRun()
        {
            try
            {
                imapService.Connect();

                Run();
            }
            catch (Exception e)
            {
                Console.WriteLine("Processing of timesheet emails failed. {0}", e);
                return;
            }
            finally
            {
                imapService.Disconnect();
            }

            Console.WriteLine("Program terminating normally.");
        }

        public void Run()
        {
            var client = imapService.client;

            var sourceFolder = client.GetFolder(sourceFolderName);

            attachmentProcessor.TryVisitFolder(sourceFolder, onCompletion(
                sourceFolder,
                client.GetFolder(processedFolderName),
                client.GetFolder(attentionFolderName)
            ));
        }
    }
}
