using System;
using System.Linq;

using MimeKit;
using MailKit;

namespace ImapAttachmentProcessing
{
    public class FileNamers
    {
        static string ReplaceFolderSeparator(string folderName, string replacement)
        {
            return folderName.Replace("/", replacement);
        }

        static string DatePart(IMessageSummary summary, string dateFormat, Func<IMessageSummary, string> missingDateLabel)
        {
            DateTimeOffset? maybeDate = summary.Envelope.Date;
            return maybeDate.HasValue ?
                maybeDate.Value.ToString(dateFormat) :
                missingDateLabel(summary);
        }

        private static readonly char[] illegalChars = System.IO.Path.GetInvalidFileNameChars();

        static string CleanFileName(string maybeDirty, string replacement = "_")
        {
            return string.Join(replacement, maybeDirty.Split(illegalChars));
        }

        static string FileNamePart(MimePart part, Func<MimePart, string> alternative) => CleanFileName(
            part.FileName != null ? part.FileName : alternative(part)
        );

        public static Func<IMailFolder, Func<MimePart, Func<IMessageSummary, string>>> FolderBasedName =
            folder => part => summary =>
            {
                var folderNamePart = ReplaceFolderSeparator(folder.FullName, "__");
                var datePart = DatePart(summary, "yyyyMM", s => string.Format("nodate-{0}", s.UniqueId));
                var fileNamePart = FileNamePart(part, p => p.ContentMd5.Substring(0, 8) + ".pdf");
                var relativeName = string.Format("{0}_{1}_{2}", folderNamePart, datePart, fileNamePart);

                return relativeName;
            };

        public static Func<MimePart, Func<IMessageSummary, string>> PartAndSummaryBasedName =
            part => summary => string.Format(
                "{0}_{1}__{2}",
                summary
                    .Envelope
                    .From
                    .Where(from => from is MailboxAddress)
                    .Select(from => (MailboxAddress)from)
                    .Select(from => from.Address)
                    .DefaultIfEmpty("unknown")
                    .First(),
                summary.InternalDate.Value.ToString("yyyy-MM-ddTHHmmss"),
                CleanFileName(part.FileName)
            );
    }
}
