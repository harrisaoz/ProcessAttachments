using System;
using System.Linq;
using System.Collections.Generic;

using MailKit.Search;
using MailKit;
using MimeKit;
using System.IO;

using Config;
using ImapAttachmentProcessing;

namespace ProcessTimesheets
{
    class TimesheetProcessor
    {
        private readonly SimpleImapService imapService;
        //private static readonly string TARGET_RECIPIENT = "timesheets@vibeteaching.co.uk";
        private readonly string targetDir;
        private readonly string sourceFolderName;
        private readonly string processedFolderName;
        private readonly string attentionFolderName;

        static void Main(string[] args)
        {
            var configLoader = new ConfigLoader();

            execute(configLoader.imapServiceSettings(),
                configLoader.imapFolderSettings(),
                configLoader.credentialsSettings(),
                configLoader.fileSystemExportSettings());
        }

        public static void execute(
            ImapService imapService,
            ImapFolders imapFolders,
            Credentials credentials,
            FileSystemExport fileSystemExport)
        {
            try
            {
                string downloadDir = fileSystemExport.ExportDirectory;
                string sourceFolder = imapFolders.SourceFolder;
                string processedFolder = imapFolders.ProcessedFolder;
                string attentionFolder = imapFolders.AttentionFolder;

                var processor = new TimesheetProcessor(imapService, credentials,
                    downloadDir,
                    sourceFolder, processedFolder, attentionFolder);
                processor.Run();
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to process required program parameters - program will now terminate. {0}");
                Console.WriteLine(e);
                return;
            }
        }

        TimesheetProcessor(ImapService serviceConfig, Credentials credentials,
            string targetDir,
            string sourceFolder, string processedFolder, string attentionFolder)
        {
            imapService = new SimpleImapService(serviceConfig, credentials);
            this.targetDir = targetDir;
            sourceFolderName = sourceFolder;
            processedFolderName = processedFolder;
            attentionFolderName = attentionFolder;
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

        public void Run()
        {
            try
            {
                var client = imapService.client;

                Console.WriteLine("Creating client");
                imapService.Connect();

                IMailFolder sourceFolder = client.GetFolder(sourceFolderName);
                var processedFolder = client.GetFolder(processedFolderName);
                var attentionFolder = client.GetFolder(attentionFolderName);
                sourceFolder.Open(FolderAccess.ReadOnly);

                //TextSearchQuery query = SearchQuery.ToContains(TARGET_RECIPIENT);
                var query = SearchQuery.DeliveredAfter(DateTime.Parse("2018-10-01"));

                var messages = sourceFolder.Search(query);
                var summaries = sourceFolder.Fetch(messages,
                    MessageSummaryItems.UniqueId |
                    MessageSummaryItems.Envelope |
                    MessageSummaryItems.BodyStructure |
                    MessageSummaryItems.InternalDate);

                //var timesheetEmails = summaries.Where(s =>
                //    s.Envelope.To.Mailboxes.Select(m => m.Address).Contains(TARGET_RECIPIENT)
                //).ToArray();
                var timesheetEmails = summaries.ToArray();

                var processed = timesheetEmails
                    .Where(e => attachmentsSaved(sourceFolder, e))
                    .ToArray();
                var requireAttention = timesheetEmails
                    .Except(processed)
                    .ToArray();

                Console.WriteLine("{0} emails were successfully processed", processed.Length);
                Console.WriteLine("{0} emails require attention", requireAttention.Length);

                onCompletion(sourceFolder, processedFolder, attentionFolder)(
                    processed.Select(summary => summary.UniqueId).ToList(),
                    requireAttention.Select(summary => summary.UniqueId).ToList()
                );
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

        bool attachmentsSaved(IMailFolder folder, IMessageSummary summary)
        {
            Console.WriteLine("Processing {0} [{1}] [{2}] [{3}]",
                summary.UniqueId,
                summary.Envelope.From,
                summary.Envelope.Subject,
                summary.InternalDate);

            if (summary.Body is BodyPartMultipart)
            {
                var attachmentBox = new AttachmentBox((BodyPartMultipart)summary.Body);
                attachmentBox.fill();
                var attachments = attachmentBox.attachmentList;

                Console.WriteLine("{0} attachments found", attachments.Count());

                var results = attachments
                    .Select(attachment => attachmentSaved(
                        folder,
                        filenameFromSummary(summary),
                        summary.UniqueId,
                        attachment)
                    ).ToArray();

                if (!results.Contains(SaveResult.Error) && results.Contains(SaveResult.Ok))
                    return true;
            }

            return false;
        }

        private string filenameFromSummary(IMessageSummary summary)
        {
            return string.Format("{0}_{1}",
                    summary.InternalDate.Value.ToString("yyyy-MM-ddTHHmmss"),
                    summary
                        .Envelope
                        .From
                        .Where(from => from is MailboxAddress)
                        .Select(from => (MailboxAddress)from)
                        .Select(from => from.Address)
                        .DefaultIfEmpty("unknown")
                        .First()
                );
        }

        enum SaveResult
        {
            Ok, Ignore, Error
        }

        private SaveResult attachmentSaved(IMailFolder folder, string filenameBase, UniqueId emailId, BodyPartBasic attachment)
        {
            var result = SaveResult.Error;

            var entity = folder.GetBodyPart(emailId, attachment);

            if (entity is MimePart)
            {
                var part = (MimePart)entity;
                if (isIgnorableAttachment(part.ContentType, part.FileName))
                {
                    logAttachment("Ignoring attachment", emailId, part.FileName, part.ContentType);
                    result = SaveResult.Ignore;
                }
                else if (isMimeTypeAcceptable(part.ContentType))
                {
                    logAttachment("MIME type ok", emailId, part.FileName, part.ContentType);

                    result = isFileDownloaded(part, filenameBase) ? SaveResult.Ok : SaveResult.Error;
                }
                else
                    logAttachment("MIME type bad", emailId, part.FileName, part.ContentType);
            }

            return result;
        }

        private bool isFileDownloaded(MimePart part, string filenameBase)
        {
            string relName = string.Format("{0}__{1}",
                filenameBase,
                FilenameCleaner.cleanFilename(part.FileName));
            string absName = Path.Combine(targetDir, relName);

            FileStream stream = null;
            try
            {
                stream = File.Open(absName, FileMode.Create);
                part.Content.DecodeTo(stream);
                Console.Write("+");
                return true;
            }
            catch (NullReferenceException nre)
            {
                Console.WriteLine("-");
                Console.WriteLine("Failed to write attachment to file. absName = {3}, stream = {0}, part = {1}. {2}", stream, part.Content, nre, absName);
            }
            catch (NotSupportedException nse)
            {
                Console.WriteLine("-");
                Console.WriteLine("The specified filename ({0}) is not valid on this filesystem. {1}", absName, nse);
            }
            catch (IOException ioe)
            {
                Console.WriteLine("-");
                Console.WriteLine("Failed to write the attachment to file. (filename {0})(part = {1})",
                    absName,
                    part);
                Console.WriteLine("{0}", ioe);
            }
            finally
            {
                try
                {
                    if (stream != null)
                        stream.Close();
                }
                catch (Exception)
                {
                    //ignore
                }
            }

            return false;
        }

        private void logAttachment(string message, UniqueId emailId, string filename, ContentType contentType)
        {
            Console.WriteLine("{0} {1}: {2} [{3}]", message, emailId, filename, contentType);
        }

        private bool isIgnorableAttachment(ContentType contentType, string name)
        {
            return contentType != null && (
                contentType.IsMimeType("text", "plain") && (name == null || name.EndsWith(".txt"))
             ||
                contentType.IsMimeType("text", "html") && (name == null || name.Contains(".htm"))
            );
        }

        private static readonly ContentType[] okAttachmentTypes =
        {
            new ContentType("image", "jpeg"),
            new ContentType("image", "png"),
            new ContentType("image", "tiff"),
            new ContentType("application", "octet-stream"),
            new ContentType("application", "pdf"),
            new ContentType("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
        };

        bool isMimeTypeAcceptable(ContentType contentType)
        {
            return okAttachmentTypes
                .Aggregate(false, (acc, c) => acc ||
                    contentType != null && contentType.IsMimeType(c.MediaType, c.MediaSubtype));
        }
    }
}
