using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Threading.Tasks;
using Azure.Core;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;

namespace SiteReview
{
    public class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            SecretClientOptions options = new SecretClientOptions()
            {
                Retry =
                {
                    Delay = TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = Azure.Core.RetryMode.Exponential
                }
            };

            var client = new SecretClient(new Uri(Globals.keyVaultUrl), new DefaultAzureCredential(), options);
            KeyVaultSecret secret = client.GetSecret(Globals.secretNameClient);
            var secretValue = secret.Value;

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(Globals.clientId)
            .WithTenantId(Globals.tenantId)
            .WithClientSecret(secretValue)
            .Build();

            var scopes = new string[] { "https://graph.microsoft.com/.default" };

            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    var authResult = await confidentialClientApplication
                    .AcquireTokenForClient(scopes)
                    .ExecuteAsync();

                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                }));

            log.LogInformation($"Created graph service client");

            return graphServiceClient;
        }

        public ClientContext appOnlyAuth(string siteUrl, ILogger log)
        {
            SecretClientOptions options = new SecretClientOptions()
            {
                Retry =
                {
                    Delay = TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = Azure.Core.RetryMode.Exponential
                }
            };

            var client = new SecretClient(new Uri(Globals.keyVaultUrl), new DefaultAzureCredential(), options);
            KeyVaultSecret secret = client.GetSecret(Globals.secretNameAppOnly);
            var secretValue = secret.Value;

            var ctx = new AuthenticationManager().GetACSAppOnlyContext(siteUrl, Globals.appOnlyId, secretValue);

            log.LogInformation($"Created app only client connection for {siteUrl}");

            return ctx;
        }

        public async Task<string> GetAccessTokenAsync()
        {
            SecretClientOptions options = new SecretClientOptions()
            {
                Retry =
                {
                    Delay = TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = Azure.Core.RetryMode.Exponential
                }
            };

            var client = new SecretClient(new Uri(Globals.keyVaultUrl), new DefaultAzureCredential(), options);
            KeyVaultSecret secret = client.GetSecret(Globals.secretNameClient);
            var secretValue = secret.Value;

            var confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(Globals.clientId)
            .WithTenantId(Globals.tenantId)
            .WithClientSecret(secretValue)
            .Build();

            var authResult = await confidentialClientApplication.AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                .ExecuteAsync();

            return authResult.AccessToken;
        }

        public class ROPCConfidentialTokenCredential : TokenCredential
        {
            string _username = "";
            string _password = "";
            string _tenantId = "";
            string _clientId = "";
            string _clientSecret = "";
            string _tokenEndpoint = "";
            ILogger _log;

            public ROPCConfidentialTokenCredential(ILogger log)
            {
                IConfiguration config = new ConfigurationBuilder()
               .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
               .AddEnvironmentVariables()
               .Build();

                _log = log;

                SecretClientOptions options = new SecretClientOptions()
                {
                    Retry =
                {
                    Delay = TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = RetryMode.Exponential
                }
                };

                var client = new SecretClient(new System.Uri(config["keyVaultUrl"]), new DefaultAzureCredential(), options);
                KeyVaultSecret secret_client = client.GetSecret(config["secretNameClient"]);
                var clientSecret = secret_client.Value;
                KeyVaultSecret secret_password = client.GetSecret(config["secretNameEmailPassword"]);
                var password = secret_password.Value;

                // Public Constructor
                _username = config["emailUserName"];
                _password = password;
                _tenantId = config["tenantId"];
                _clientId = config["clientId"];
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
}
