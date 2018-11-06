﻿using System;
using System.Linq;
using System.Collections.Generic;

using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.IO;

namespace ProcessAttachments
{
    class InvoiceDownloader
    {
        ImapClient client;
        // Required parameters
        string baseDirectory = "";
        string username = "";
        string password = "";
        string imapHostname = "";
        string[] folders = null;

        public static void execute(string[] args)
        {
            try
            {
                string hostname = args[0];
                string baseDir = args[1];
                string user = args[2];
                string pass = args[3];
                string[] folders = args.Skip(4).ToArray();

                var downloader = new InvoiceDownloader(hostname, baseDir, user, pass, folders);
                downloader.run();
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to process required program parameters - program will now terminate. {0}");
                Console.WriteLine(e);
                return;
            }
        }
        public InvoiceDownloader(string imapHostname, string baseDirectory, string username, string password, string[] folders)
        {
            this.imapHostname = imapHostname;
            this.baseDirectory = baseDirectory;
            this.username = username;
            this.password = password;
            this.folders = folders;
        }

        public void run()
        {
            try
            {
                Console.WriteLine("Creating client");
                client = new ImapClient();
                ulong startingModSeq = 0L;

                foreach (var folder in openFolders(client))
                {
                    uint totalSize = traverseFolder(folder, 0, startingModSeq);
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
                try
                {
                    client.Disconnect(true);
                    client.Dispose();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unexpected failure to disconnect the imap client. {0}", e);
                    //ignore, we just want to clean up.
                }
            }

            Console.WriteLine("Program terminating normally.");
        }

        IEnumerable<IMailFolder> openFolders(ImapClient client)
        {
            Console.WriteLine("Opening mailbox");
            client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            client.Connect(imapHostname, 993, true);
            client.Authenticate(username, password);

            var topFolder = client.GetFolder(client.PersonalNamespaces[0]);

            Console.WriteLine("Filtering folders against list: {0}", folders);
            foreach (var folder in topFolder.GetSubfolders(false))
            {
                Console.WriteLine("Checking folder {0}", folder.Name);
                if (Array.Exists<string>(folders, name => name.ToUpper().Equals(folder.Name.ToUpper())))
                {
                    Console.WriteLine("Found matching folder: {0}", folder.Name);
                    yield return folder;
                }
            }
        }

        private uint traverseFolder(IMailFolder f0, uint visitedSize, ulong modSeq)
        {
            uint size = visitFolderAction(f0, modSeq);

            foreach (var f1 in f0.GetSubfolders(false))
            {
                size = traverseFolder(f1, size, modSeq);
            }

            return visitedSize + size;
        }

        private uint visitFolderAction(IMailFolder folder, ulong modSeq)
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
            return maybeDate.HasValue ? maybeDate.Value.ToString("yyyyMM") : String.Format("nodate-{0}", summary.UniqueId);
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

                                string invoiceLabel = String.Format("{0}_{1}_{2}", serviceLabel, dateLabel, filename);

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
                                catch (System.NullReferenceException nre)
                                {
                                    Console.WriteLine();
                                    Console.WriteLine("Failed to write attachment to file. absName = {3}, stream = {0}, part = {1}. {2}", stream, part.Content, nre, absName);
                                }
                                catch (System.NotSupportedException nse)
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
