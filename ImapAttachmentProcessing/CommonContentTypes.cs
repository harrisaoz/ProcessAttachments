using MimeKit;

namespace ImapAttachmentProcessing
{
    public class CommonContentTypes
    {
        static ContentType ct(string mediaType, string subType)
        {
            return new ContentType(mediaType, subType);
        }

        public readonly static ContentType
            IMAGE_JPEG = ct("image", "jpeg"),
            IMAGE_JPG = ct("image", "jpg"), // not an official MIME type, but some (broken) software apparently uses this!
            IMAGE_PNG = ct("image", "png"),
            IMAGE_TIFF = ct("image", "tiff"),
            WORD_DOC = ct(
                "application",
                "vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            TEXT_PLAIN = ct("text", "plain"),
            TEXT_HTML = ct("text", "html"),
            PDF_NORMAL = ct("application", "pdf"),
            PDF_GENERAL = ct("application", "octet-stream");
    }
}
