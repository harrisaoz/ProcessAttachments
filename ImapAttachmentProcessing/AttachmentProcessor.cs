using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using MimeKit;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;

namespace ImapAttachmentProcessing
{
    public class AttachmentProcessor
    {
        private readonly IEnumerable<ContentType> attachmentTypes;
        private readonly Func<IMailFolder, Func<MimePart, Func<IMessageSummary, string>>> fileNamer;
        private readonly AttachmentDownloader downloader;
        private readonly Func<ContentType, string, bool> isIgnoreableAttachment;
        private readonly Func<IMailFolder, IList<IMessageSummary>> fetch;

        public AttachmentProcessor(IEnumerable<ContentType> attachmentTypesToProcess,
            Func<IMailFolder, Func<MimePart, Func<IMessageSummary, string>>> fileNamer,
            AttachmentDownloader downloader,
            Func<ContentType, string, bool> isIgnorableAttachment,
            SearchQuery searchQuery,
            MessageSummaryItems fetchFlags
        )
        {
            this.attachmentTypes = attachmentTypesToProcess;
            this.fileNamer = fileNamer;
            this.downloader = downloader;
            this.isIgnoreableAttachment = isIgnorableAttachment;

            if (searchQuery != null)
            {
                fetch = folder =>
                {
                    var messages = folder.Search(searchQuery);
                    return folder.Fetch(messages, fetchFlags);
                };
            }
            else
            {
                fetch = folder => folder.Fetch(0, -1, 0, fetchFlags);
            }
        }

        public bool IsMimeTypeAcceptable(ContentType contentType) => attachmentTypes
                .Aggregate(false, (acc, c) => acc ||
                    contentType != null && contentType.IsMimeType(c.MediaType, c.MediaSubtype));

        public int TryVisitFolder(
            IMailFolder folder,
            Action<List<UniqueId>, List<UniqueId>> onCompletion)
        {
            try
            {
                return VisitFolder(folder, onCompletion);
            }
            catch (ImapCommandException commandException)
            {
                Console.WriteLine("Failed to open folder {0}. {1}", folder.Name, commandException.Message);
                return 0;
            }
            finally
            {
                try { folder.Close(); }
                catch (Exception) { }
                // ignore - it's ok if the folder was already closed; we just want to ensure that it is closed.
            }
        }

        private int VisitFolder(IMailFolder folder,
            Action<List<UniqueId>, List<UniqueId>> onCompletion)
        {
            folder.Open(FolderAccess.ReadOnly);

            var summaries = fetch(folder).ToArray();

            var processed = summaries
                .Where(summary => saveEmailAttachments(folder, summary)
                    .Equals(SaveResult.Ok)
                )
                .ToArray();

            var requireAttention = summaries
                .Except(processed)
                .ToArray();

            Console.WriteLine("{0} emails were successfully processed", processed.Length);
            Console.WriteLine("{0} emails require attention", requireAttention.Length);

            onCompletion(
                processed.Select(summary => summary.UniqueId).ToList(),
                requireAttention.Select(summary => summary.UniqueId).ToList()
            );

            return processed.Count();
        }

        private SaveResult saveEmailAttachments(IMailFolder folder, IMessageSummary summary)
        {
            if (summary.Body is BodyPartMultipart body)
            {
                var attachmentBox = new AttachmentBox(body);
                attachmentBox.fill();

                var attachments = attachmentBox.attachmentList;

                var results = attachments
                    .Select(attachment => saveAttachment(
                        folder, summary, attachment
                    )).ToArray();

                if (!results.Contains(SaveResult.Error) && results.Contains(SaveResult.Ok))
                    return SaveResult.Ok;
                else if (results.Contains(SaveResult.Error))
                    return SaveResult.Error;
            }

            return SaveResult.Ignore;
        }

        public SaveResult saveAttachment(IMailFolder folder, IMessageSummary summary, BodyPartBasic attachment)
        {
            var result = SaveResult.Error;

            var entity = folder.GetBodyPart(summary.UniqueId, attachment);

            if (entity is MimePart)
            {
                var part = (MimePart)entity;
                var log = logAttachment(summary.UniqueId, part.FileName, part.ContentType);

                if (isIgnoreableAttachment(part.ContentType, part.FileName))
                {
                    log("Ignoring attachment");
                    result = SaveResult.Ignore;
                }
                else if (IsMimeTypeAcceptable(part.ContentType))
                {
                    log("MIME type ok");

                    result = downloader.downloadAttachment(
                        part.Content, this.fileNamer(folder)(part)(summary)
                    );
                }
                else
                    log("MIME type bad");
            }

            return result;
        }

        private Func<UniqueId, string, ContentType, Action<string>> logAttachment =
            (emailId, fileName, contentType) => message =>
            {
                Console.WriteLine("{0} {1}: {2} [{3}]", message, emailId, fileName, contentType);
            };
    }
}
