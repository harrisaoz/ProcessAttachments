using System.Collections.Generic;

using System.Linq;

using MailKit;

namespace ImapAttachmentProcessing
{
    public class AttachmentBox
    {
        private readonly BodyPartMultipart top;
        public readonly List<BodyPartBasic> attachmentList;

        public AttachmentBox(BodyPartMultipart multipart)
        {
            top = multipart;
            attachmentList = new List<BodyPartBasic>();
        }

        public void fill()
        {
            fill(top);
        }

        public void fill(BodyPartMultipart part)
        {
            attachmentList.AddRange(
                part
                    .BodyParts
                    .OfType<BodyPartBasic>()
            );

            foreach (var mp in part.BodyParts.Where(p => p is BodyPartMultipart))
            {
                fill((BodyPartMultipart)mp);
            }
        }
    }
}
