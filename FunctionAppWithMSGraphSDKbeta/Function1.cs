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
using Azure.Core;
using System.Threading;
using System.Net.Http;
using System.Collections.Generic;

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


            // For this sample to work, your application should have the delegated permissions added and consented by an admin:
            //  "Directory.Read.All"
            //  "AuditLog.Read.All"


            // The Username-Password flow requires that you request the previously
            // consented scopes, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "AuditLog.Read.All", "Directory.ReadWrite.All" };


            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "{put your tenant id here}";

            // Values from app registration
            var clientId = "{put the client id for your application here}";
            var clientSecret = "{put your application's client secret here}";

            var username = "{replace with the appropriate username}";
            var password = "{replace with the password}";


            // Use our custom TokenCredential class: ROPCConfidentialTokenCredential
            var ROPCCredential = new ROPCConfidentialTokenCredential(username, password, tenantId, clientId, clientSecret);


            var graphClient = new GraphServiceClient(ROPCCredential, scopes);

            var users = await graphClient.Users.Request().Select("signInActivity").GetAsync();

            var firstSignInDateTime = users.CurrentPage[0].SignInActivity.LastSignInDateTime.ToString();

            /*
             * End Microsoft Graph SDK related code
             */


            string responseMessage = firstSignInDateTime;

            return new OkObjectResult(responseMessage);
        }
    }

    public class ROPCConfidentialTokenCredential : Azure.Core.TokenCredential
    {
        // Implementation of the Azure.Core.TokenCredential class
        string _username = "";
        string _password = "";
        string _tenantId = "";
        string _clientId = "";
        string _clientSecret = "";

        string _tokenEndpoint = "";

        public ROPCConfidentialTokenCredential(string username, string password, string tenantId, string clientId, string clientSecret)
        {
            // Public Constructor
            _username = username;
            _password = password;
            _tenantId = tenantId;
            _clientId = clientId;
            _clientSecret = clientSecret;

            _tokenEndpoint = "https://login.microsoftonline.com/" + _tenantId + "/oauth2/v2.0/token";
        }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            HttpClient httpClient = new HttpClient();

            // Create the request body
            var Parameters = new List<KeyValuePair<string, string>> 
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                new KeyValuePair<string, string>("username", _username),
                new KeyValuePair<string, string>("password", _password),
                new KeyValuePair<string, string>("grant_type", "password")
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint) 
            {
                Content = new FormUrlEncodedContent(Parameters)
            };
            var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
            dynamic responseJson = JsonConvert.DeserializeObject(response);
            var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
            return new AccessToken(responseJson.access_token.ToString(), expirationDate);
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            HttpClient httpClient = new HttpClient();

            // Create the request body
            var Parameters = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                new KeyValuePair<string, string>("username", _username),
                new KeyValuePair<string, string>("password", _password),
                new KeyValuePair<string, string>("grant_type", "password")
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
            {
                Content = new FormUrlEncodedContent(Parameters)
            };
            var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
            dynamic responseJson = JsonConvert.DeserializeObject(response);
            var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
            return new ValueTask<AccessToken>(new AccessToken(responseJson.access_token.ToString(), expirationDate));
        }
    }
}
