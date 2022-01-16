using System;
using System.Collections.Generic;
using System.Linq;

using MailKit;
using MailKit.Net.Imap;

using Config;

namespace ImapAttachmentProcessing
{
    public class SimpleImapService
    {
        public readonly ImapClient client;
        private readonly ImapService serviceConfig;
        private readonly Credentials credentials;

        public SimpleImapService(ImapService imapService, Credentials credentials)
        {
            client = new ImapClient();
            serviceConfig = imapService;
            this.credentials = credentials;
        }

        public void Connect()
        {
            Console.WriteLine("Opening mailbox");
            client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            client.Connect(serviceConfig.Hostname, serviceConfig.Port, true);
            client.Authenticate(credentials.Username, credentials.Password);
        }

        public void Disconnect()
        {
            try
            {
                client.Disconnect(true);
                client.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine("Unexpected failure to disconnect the imap client. {0}", e);
            }
        }
    }
}
