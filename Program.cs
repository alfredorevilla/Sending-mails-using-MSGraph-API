using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace Sending_mails_using_Graph_API
{
    class Program
    {
        private static readonly string clientId = "<value>";
        private static readonly string tenantID = "<value>";
        private static readonly string clientSecret = "<value>";

        static async Task Main(string[] args)
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantID)
                .WithClientSecret(clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var message = new Message
            {
                Subject = "Meet for lunch?",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = "The new cafeteria is open."
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "<value>";
                        }
                    }
                },
                CcRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "<value>";
                        }
                    }
                }
            };

            var saveToSentItems = true;
            await graphClient
                .Users["shared1@alfredorevillamsft1.OnMicrosoft.com"]
              .SendMail(message, saveToSentItems)
              .Request()
              .PostAsync();

            Console.WriteLine(0);

        }
    }
}
