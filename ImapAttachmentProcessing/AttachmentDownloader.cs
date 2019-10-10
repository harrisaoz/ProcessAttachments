using System;
using System.IO;

using MimeKit;

namespace ImapAttachmentProcessing
{
    public class AttachmentDownloader
    {
        private readonly string downloadFolder;

        public AttachmentDownloader(string downloadFolder) {
            this.downloadFolder = downloadFolder;
        }

        Func<string, IMimeContent, Action<string>> log0 = (fileName, content) => formatText => {
            Console.WriteLine(formatText, fileName, content);
        };

        Func<string, IMimeContent, Action<string, Exception>> logE0 = (fileName, content) => (formatText, exception) => {
            Console.WriteLine(formatText, fileName, content, exception);
        };

        public SaveResult downloadAttachment(IMimeContent content, string fileName)
        {
            FileStream stream = null;
            var log = log0(fileName, content);
            var logE = logE0(fileName, content);

            var absoluteFileName = Path.Combine(
                assertDirectoryExists(Path.Combine(this.downloadFolder)),
                fileName);
            try
            {
                stream = File.Open(absoluteFileName, FileMode.CreateNew);
                content.DecodeTo(stream);
                log("+");
                return SaveResult.Ok;
            }
            catch (NullReferenceException nre)
            {
                log("-");
                logE("Failed to write attachment to file. (filename = {0}) {2}", nre);
            }
            catch (NotSupportedException nse)
            {
                log("-");
                logE("The specified filename ({0}) is not valid on this filesystem. {2}", nse);
            }
            catch (IOException ioe)
            {
                log("="); // file already exists
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

            return SaveResult.Error;
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
