using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System.Threading.Tasks;

namespace SendDraftEmail
{
    class Mailer
    {
        public static async Task SendDraftEmailAsync(string draftMessageId)
        {
            // Set up the GraphServiceClient with authentication
            var graphClient = GetGraphServiceClient();

            // Get the draft message by ID
            var draftMessage = await graphClient.Me.Messages[draftMessageId]
                .Request()
                .GetAsync();

            // Create a new message object to send
            var messageToSend = new Microsoft.Graph.Message
            {
                Subject = draftMessage.Subject,
                Body = draftMessage.Body,
                ToRecipients = draftMessage.ToRecipients,
                CcRecipients = draftMessage.CcRecipients,
                BccRecipients = draftMessage.BccRecipients,
                Importance = draftMessage.Importance
            };

            // Send the message
            await graphClient.Me.SendMail(messageToSend, true)
                .Request()
                .PostAsync();
        }

        static GraphServiceClient GetGraphServiceClient()
        {
            // Set up the authentication provider
            /*            var clientId = "ba08ca8a-5ec7-47f2-bc3d-216db4ee09da";
                        var tenantId = "b958b0f9-abfa-4de1-b635-451887da84eb";
                        var clientSecret = "iP48Q~YPkEVfVn6GUEP_fVTDnCW1V8cQJ5.KpcPO";*/
            // test
            var clientId = "4d7c4a81-b69b-4889-9dbc-dad9830cf421";
            var tenantId = "f8cdef31-a31e-4b4a-93e4-5f571e91255a";
            var clientSecret = "4D-8Q~jtXDPJ9IiQbIxudtc-ssz5Vy8dBfjl5b8R";

            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();

            var authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Set up the GraphServiceClient with authentication
            var graphClient = new GraphServiceClient(authProvider);
            return graphClient;
        }
    }
}
