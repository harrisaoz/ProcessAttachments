using System;
using System.Collections.Generic;

using MailKit;
using MailKit.Net.Imap;

namespace ImapAttachmentProcessing
{
    using GetFolderByName = Func<string, IEnumerable<IMailFolder>>;

    public interface IMailFolderFinder
    {
        IEnumerable<IMailFolder> FindByName(string folderName);
    }

    public class MailFolderFinder : IMailFolderFinder
    {
        private readonly GetFolderByName getFolderByName;

        public MailFolderFinder(
                Func<ImapClient, Func<string, IEnumerable<IMailFolder>>> findByName,
                ImapClient client)
        {
            this.getFolderByName = findByName(client);
        }

        public IEnumerable<IMailFolder> FindByName(string folderName) =>
            this.getFolderByName(folderName);
    }
}