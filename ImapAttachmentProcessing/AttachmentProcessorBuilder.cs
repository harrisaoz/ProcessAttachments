using System;
using System.Collections.Generic;
using System.Linq;

using MimeKit;
using MailKit;
using MailKit.Search;

namespace ImapAttachmentProcessing
{
    public class AttachmentProcessorBuilder
    {
        private IList<ContentType> attachmentTypes;
        private string downloadFolder;
        private Func<IMailFolder, Func<MimePart, Func<IMessageSummary, string>>> fileNamer;
        private Func<ContentType, string, bool> isIgnorableAttachment;
        private SearchQuery searchQuery;
        private MessageSummaryItems fetchFlags;

        public AttachmentProcessorBuilder()
        {
            attachmentTypes = new List<ContentType>();
            downloadFolder = "";
            fileNamer = folder => part => summary => "";
            isIgnorableAttachment = (contentType, fileNamer) => false;
            searchQuery = null;
            fetchFlags = MessageSummaryItems.None;
        }

        public AttachmentProcessorBuilder PermitAttachmentType(ContentType attachmentType)
        {
            attachmentTypes.Add(attachmentType);
            return this;
        }

        public AttachmentProcessorBuilder DownloadTo(string folder)
        {
            downloadFolder = folder;
            return this;
        }

        public AttachmentProcessorBuilder NameDownloadsByPartAndSummary(
            Func<MimePart, Func<IMessageSummary, string>> fileNamer)
        {
            this.fileNamer = folder => fileNamer;
            return this;
        }

        public AttachmentProcessorBuilder NameDownloadsByFolder(Func<IMailFolder, Func<MimePart, Func<IMessageSummary, string>>> fileNamer)
        {
            this.fileNamer = fileNamer;
            return this;
        }

        public AttachmentProcessorBuilder IgnoringAttachments(Func<ContentType, string, bool> isIgnorable)
        {
            this.isIgnorableAttachment = isIgnorable;
            return this;
        }

        public AttachmentProcessorBuilder UsingSearchQuery(SearchQuery searchQuery)
        {
            this.searchQuery = searchQuery;
            return this;
        }

        public AttachmentProcessorBuilder IncludeMessageHeaders(MessageSummaryItems fetchFlags)
        {
            this.fetchFlags = fetchFlags;
            return this;
        }

        public AttachmentProcessor Build()
        {
            return new AttachmentProcessor(
                attachmentTypes,
                fileNamer,
                new AttachmentDownloader(downloadFolder),
                isIgnorableAttachment,
                searchQuery,
                fetchFlags);
        }
    }
}
