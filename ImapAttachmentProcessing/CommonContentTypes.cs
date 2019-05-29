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
            JPEG = ct("image", "jpeg"),
            PNG = ct("image", "png"),
            TIFF = ct("image", "tiff"),
            WORD_DOC = ct(
                "application",
                "vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            TXT = ct("text", "plain"),
            HTML = ct("text", "html"),
            PDF_NORMAL = ct("application", "pdf"),
            PDF_GENERAL = ct("application", "octet-stream");

        public readonly static ContentType[]
            IMAGES = {
                JPEG,
                PNG,
                TIFF
            },
            PDF = {
                PDF_NORMAL,
                PDF_GENERAL
            },
            TEXT = {
                TXT,
                HTML
            };

    }
}
