using System;
using System.Collections.Generic;

using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.IO;

namespace ProcessAttachments
{
    class Program
    {
        static ImapClient client;
        static string baseDirectory = Path.Combine("D:", "email-attachments", "Inbox");

        static void Main(string[] args)
        {
            try
            {
                client = new ImapClient();
                ulong startingModSeq = ulong.Parse(args[0]);
                var folder = openFolder(client);

                uint totalSize = traverseFolder(folder, 0, startingModSeq);
                Console.WriteLine("Total size of items with attachments: {0:N} bytes", totalSize);
            }
            catch (Exception e)
            {
                Console.WriteLine("A starting modification sequence is required. {0}", e);
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
                    //ignore, we just want to clean up.
                }
            }
        }

        static IMailFolder openFolder(ImapClient client)
        {
            client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            client.Connect("imap.gmail.com", 993, true);
            client.Authenticate("user", "password");

            //var folder = client.GetFolder(SpecialFolder.All);
            var folder = client.Inbox;
            folder.Open(FolderAccess.ReadOnly);

            return folder;
        }

        static uint traverseFolder(IMailFolder f0, uint visitedSize, ulong modSeq)
        {
            uint size = visitFolderAction(f0, modSeq);

            foreach (var f1 in f0.GetSubfolders(false))
            {
                size = traverseFolder(f1, size, modSeq);
            }

            return visitedSize + size;
        }

        private static uint visitFolderAction(IMailFolder folder, ulong modSeq)
        {
            uint folderSize = 0;
            try
            {
                folder.Open(FolderAccess.ReadOnly);

                Console.WriteLine("Visiting folder {0} [{1,5}]", folder.FullName, folder.Count);

                foreach (var summary in folder.Fetch(0, -1, 0,
                    MessageSummaryItems.UniqueId |
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
                catch (FolderNotOpenException notOpen)
                {
                    // ignore failures to close folders that were never opened.
                }
            }

            Console.WriteLine();
            Console.WriteLine("Visited folder {0} [{1,16:N}]", folder.FullName, folderSize);
            return folderSize;
        }

        private static uint visitMessageAction(IMailFolder folder, IMessageSummary summary)
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

                            var dir = assertDirectoryExists(Path.Combine(baseDirectory, folder.FullName.Replace("/", "__"), summary.UniqueId.ToString()));
                            var part = (MimePart)entity;
                            var relName = (part.FileName == null ? part.ContentMd5.Substring(0, 8) : part.FileName).Replace(':', '_');
                            var absName = Path.Combine(dir, relName);

                            FileStream stream = null;
                            try
                            {
                                stream = File.Open(absName, FileMode.CreateNew);
                                part.Content.DecodeTo(stream);
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
                            catch (IOException ioe)
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
                                catch (Exception e)
                                {
                                    //ignore
                                }
                            }

                            Console.Write("-");
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
                return visitMessageAction(openFolder(client), summary);
            }
        }

        static string assertDirectoryExists(string dir)
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            return dir;
        }
    }
}
