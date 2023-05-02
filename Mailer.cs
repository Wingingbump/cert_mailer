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
            // test
            var clientId = "";
            var tenantId = "";
            var clientSecret = "";

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
