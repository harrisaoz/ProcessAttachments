using System;
using System.Linq;
using System.Collections.Generic;

using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.IO;

using Config;
using ImapAttachmentProcessing;

namespace DownloadInvoices
{
    class InvoiceDownloader
    {
        readonly private SimpleImapService imapService;
        readonly string baseDirectory;
        readonly string sourceFolderName;

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
                string baseDir = fileSystemExport.ExportDirectory;
                string folder = imapFolders.SourceFolder;

                var downloader = new InvoiceDownloader(imapService, credentials, baseDir, folder);
                downloader.run();
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to process required program parameters - program will now terminate. {0}");
                Console.WriteLine(e);
                return;
            }
        }

        public InvoiceDownloader(
            ImapService serviceConfig,
            Credentials credentials,
            string baseDirectory,
            string folder)
        {
            this.baseDirectory = baseDirectory;
            this.imapService = new SimpleImapService(serviceConfig, credentials);
            sourceFolderName = folder;
        }

        public void run()
        {
            try
            {
                imapService.Connect();

                foreach (var folder in openFolders(imapService.client))
                {
                    uint totalSize = TraverseFolder(folder, 0);
                    Console.WriteLine("Total size of items with attachments: {0:N} bytes", totalSize);
                }
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

        IEnumerable<IMailFolder> openFolders(ImapClient client)
        {
            Console.WriteLine("Opening mailbox");

            var topFolder = client.GetFolder(client.PersonalNamespaces[0]);

            return imapService.FilterByName(topFolder, new[] { this.sourceFolderName });
        }

        private uint TraverseFolder(IMailFolder f0, uint visitedSize)
        {
            uint size = VisitFolderAction(f0);

            foreach (var f1 in f0.GetSubfolders(false))
            {
                size = TraverseFolder(f1, size);
            }

            return visitedSize + size;
        }

        private uint VisitFolderAction(IMailFolder folder)
        {
            uint folderSize = 0;
            try
            {
                folder.Open(FolderAccess.ReadOnly);

                Console.WriteLine("Visiting folder {0} [{1,5}]", folder.FullName, folder.Count);

                foreach (var summary in folder.Fetch(0, -1, 0,
                    MessageSummaryItems.UniqueId |
                    MessageSummaryItems.Envelope |
                    MessageSummaryItems.GMailLabels |
                    MessageSummaryItems.BodyStructure |
                    MessageSummaryItems.Size |
                    MessageSummaryItems.ModSeq))
                {
                    folderSize += visitMessageAction(folder, summary);
                }
            }
            catch (ImapCommandException commandException)
            {
                Console.WriteLine("Failed to open folder {0}. {1}", folder.Name, commandException.Message);
            }
            finally
            {
                try
                {
                    folder.Close();
                }
                catch (FolderNotOpenException)
                {
                    // ignore failures to close folders that were never opened.
                }
            }

            Console.WriteLine();
            Console.WriteLine("Visited folder {0} [{1,16:N}]", folder.FullName, folderSize);
            return folderSize;
        }

        string makeDateLabel(IMessageSummary summary)
        {
            DateTimeOffset? maybeDate = summary.Envelope.Date;
            return maybeDate.HasValue ? maybeDate.Value.ToString("yyyyMM") : string.Format("nodate-{0}", summary.UniqueId);
        }

        uint visitMessageAction(IMailFolder folder, IMessageSummary summary)
        {
            try
            {
                if (summary.Body is BodyPartMultipart)
                {
                    bool hasMimeAttachments = false;
                    //Console.WriteLine("Found multi-part body for message {0}", summary.UniqueId);
                    foreach (var attachment in summary.Attachments)
                    {
                        var entity = folder.GetBodyPart(summary.UniqueId, attachment);

                        if (entity is MimePart)
                        {
                            hasMimeAttachments = true;
                            Console.Write(".");
                            var part = (MimePart)entity;

                            if (part.ContentType.IsMimeType("application", "pdf") || part.ContentType.IsMimeType("application", "octet-stream"))
                            {
                                string serviceLabel = folder.FullName.Replace("/", "__");
                                string dateLabel = makeDateLabel(summary);
                                string filename = (part.FileName == null ? part.ContentMd5.Substring(0, 8) + ".pdf" : part.FileName).Replace(':', '_');
                                if (filename.Contains("detalhe") || filename.EndsWith(".zip"))
                                {
                                    continue;
                                }

                                string invoiceLabel = string.Format("{0}_{1}_{2}", serviceLabel, dateLabel, filename);

                                Console.WriteLine("{0}: [MIME type: {1}]", invoiceLabel, part.ContentType.MimeType);

                                var dir = assertDirectoryExists(Path.Combine(baseDirectory));
                                var absName = Path.Combine(dir, invoiceLabel);

                                FileStream stream = null;
                                try
                                {
                                    stream = File.Open(absName, FileMode.CreateNew);
                                    part.Content.DecodeTo(stream);
                                    Console.Write("-");
                                }
                                catch (NullReferenceException nre)
                                {
                                    Console.WriteLine();
                                    Console.WriteLine("Failed to write attachment to file. absName = {3}, stream = {0}, part = {1}. {2}", stream, part.Content, nre, absName);
                                }
                                catch (NotSupportedException nse)
                                {
                                    Console.WriteLine();
                                    Console.WriteLine("The specified filename ({0}) is not valid on this filesystem. {1}", absName, nse);
                                }
                                catch (IOException)
                                {
                                    Console.Write("=");
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
                            }
                        }
                    }
                    if (hasMimeAttachments)
                        return summary.Size.GetValueOrDefault(0);
                }
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception occurred while visiting message {1} in folder {0}. {2}", folder.FullName, summary.ModSeq, e);
                throw e;
            }
        }

        string assertDirectoryExists(string dir)
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            return dir;
        }
    }
}
