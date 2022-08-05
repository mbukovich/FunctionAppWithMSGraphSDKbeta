using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Azure.Identity;
using Microsoft.Graph;

namespace FunctionAppWithMSGraphSDKbeta
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            /*
             * Begin Microsoft Graph SDK related code
             */


            // For this sample to work, your application should have the application permissions:
            //  "Directory.Read.All"
            //  "AuditLog.Read.All"


            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "{put your tenant id here}";

            // Values from app registration
            var clientId = "{put the client id for your application here}";
            var clientSecret = "{put your application's client secret here}";

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            var users = await graphClient.Users.Request().Select("signInActivity").GetAsync();

            var firstSignInDateTime = users.CurrentPage[0].SignInActivity.LastSignInDateTime.ToString();

            /*
             * End Microsoft Graph SDK related code
             */


            string responseMessage = firstSignInDateTime;

            return new OkObjectResult(responseMessage);
        }
    }
}
