using System;
using System.Collections.Generic;
using System.Text;

using System.Linq;

using MailKit;

namespace ProcessAttachments
{

    class AttachmentBox
    {
        private BodyPartMultipart top;
        public readonly List<BodyPartBasic> attachmentList;

        public AttachmentBox(BodyPartMultipart multipart)
        {
            this.top = multipart;
            this.attachmentList = new List<BodyPartBasic>();
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
                .Where(p => p is BodyPartBasic)
                .Select(b => (BodyPartBasic)b)
            );

            foreach (var mp in part.BodyParts.Where(p => p is BodyPartMultipart))
            {
                fill((BodyPartMultipart)mp);
            }
        }
    }
}
