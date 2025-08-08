using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Graph.Models;
using SiteReview;
using static SiteReview.Auth;
using System.IO;
using Microsoft.Graph.Sites;
using Microsoft.Graph.Users.Item.SendMail;

namespace SiteReviewProB
{
    public class SiteReviewProB
    {
        [Function("SiteReviewProB")]
        public static async Task Run(
        [TimerTrigger("0 0 * * * *")] TimerInfo myTimer, ILogger log, FunctionContext executionContext)
        {
            log.LogInformation($"SiteReviewProB timer trigger function executed at: {DateTime.Now}");

            try
            {
                var basePath = Environment.GetEnvironmentVariable("AzureFunctionsJobRoot") ?? Directory.GetCurrentDirectory();

                var config = new ConfigurationBuilder()
                    .SetBasePath(basePath)
                    .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                    .AddEnvironmentVariables()
                    .Build();

                var graphAPIAuth = new Auth().graphAuth(log);
                log.LogInformation("Graph API authentication successful.");

                var sites = await GetProtectedBSites(graphAPIAuth, log);
                log.LogInformation($"Retrieved {sites.Count} protected B sites.");

                var publicSites = new List<Site>();

                foreach (var site in sites)
                {
                    log.LogInformation($"Processing site: {site.DisplayName}");

                    var sitePrivacySetting = await GetSitePrivacySetting(graphAPIAuth, site, log);
                    log.LogInformation($"Site {site.DisplayName} privacy setting: {sitePrivacySetting}");

                    if (sitePrivacySetting == "Public")
                    {
                        publicSites.Add(site);
                        log.LogInformation($"Site {site.DisplayName} added to public sites list.");
                    }
                }

                if (publicSites.Any())
                {
                    log.LogInformation("Public sites found, sending report email.");
                    var emailRecipients = config["AdminEmails"];
                    log.LogInformation($"AdminEmails: {emailRecipients}");
                    if (string.IsNullOrEmpty(emailRecipients))
                    {
                        log.LogError("AdminEmails setting is not configured.");
                        throw new InvalidOperationException("AdminEmails setting is not configured.");
                    }

                    var recipientEmails = emailRecipients.Split(',', StringSplitOptions.RemoveEmptyEntries);

                    var scopes = new[] { "user.read", "mail.send" };
                    var auth = new ROPCConfidentialTokenCredential(log);
                    var graphClient = new GraphServiceClient(auth, scopes);

                    await ProtectedBEmail(recipientEmails, publicSites, graphClient, log);
                }

                log.LogInformation("Function app executed successfully");
            }
            catch (Exception ex)
            {
                log.LogError($"An error occurred: {ex.Message}");
            }
            log.LogInformation("Function app pro b executed successfully");
            //  return new OkObjectResult("HTTP trigger executed successfully.");
        }

        private static async Task<List<Site>> GetProtectedBSites(GraphServiceClient graphClient, ILogger log)
        {
            var sites = new List<Site>();
            log.LogInformation("Getting Protected B Sites.");

            try
            {
                var response = await graphClient.Sites.GetAsync(requestConfig =>
                {
                    requestConfig.Headers.Add("ConsistencyLevel", "eventual");
                });

                while (response != null)
                {
                    var currentSites = response.Value;

                    log.LogInformation($"Processing {currentSites.Count} sites from current page.");

                    foreach (var site in currentSites)
                    {
                        var group = await Common.GetGroupFromSite(site, graphClient, log);
                        if ((group?.AssignedLabels != null && group.AssignedLabels.Any(label => label.DisplayName.Contains("Protected B"))) ||
                            site.WebUrl.Contains("/teams/b"))
                        {
                            log.LogInformation($"Site {site.DisplayName} classified as Protected B.");
                            sites.Add(site);
                        }
                    }

                    if (!string.IsNullOrEmpty(response.OdataNextLink))
                    {
                        var nextRequestBuilder = new SitesRequestBuilder(response.OdataNextLink, graphClient.RequestAdapter);
                        response = await nextRequestBuilder.GetAsync();
                    }
                    else
                    {
                        response = null;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Failed to get Protected B sites: {ex.Message}");
            }

            log.LogInformation($"Total Protected B Sites retrieved: {sites.Count}");
            return sites;
        }

        private static async Task<string> GetSitePrivacySetting(GraphServiceClient graphClient, Site site, ILogger log)
        {
            try
            {
                log.LogInformation($"Requesting privacy setting for site: {site.DisplayName}");

                var group = await Common.GetGroupFromSite(site, graphClient, log);
                if (group != null)
                {
                    log.LogInformation($"Group {group.Id} has visibility: {group.Visibility}");
                    return group.Visibility?.ToString();
                }
                else
                {
                    log.LogWarning($"No associated group found for site: {site.DisplayName}");
                }
            }
            catch (Exception ex)
            {
                log.LogError($"An error occurred while getting the site privacy setting for site {site.DisplayName}: {ex.Message}");
            }

            return "Unknown";
        }

        private static async Task ProtectedBEmail(string[] recipientEmails, List<Site> publicSites, GraphServiceClient graphClient, ILogger log)
        {
            var emailContent = ProtectedBEmailContent(publicSites);
            var emailMessage = new Message
            {
                Subject = "ProtectedB Public Sites Report",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = emailContent
                },
                ToRecipients = recipientEmails.Select(email => new Recipient { EmailAddress = new EmailAddress { Address = email } }).ToList()
            };

            try
            {
                var requestBody = new SendMailPostRequestBody
                {
                    Message = emailMessage,
                    SaveToSentItems = true
                };

                await graphClient.Users[Globals.emailUserName].SendMail.PostAsync(requestBody);

                log.LogInformation("Report email sent successfully.");
            }
            catch (Exception ex)
            {
                log.LogError($"Failed to send email: {ex.Message}");
            }
        }

        private static string ProtectedBEmailContent(List<Site> publicSites)
        {
            return "The following ProtectedB sites are public:\n" + string.Join("\n", publicSites.Select(site => $"{site.DisplayName}: {site.WebUrl}"));
        }
    }
}