using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using SiteReview;
using static SiteReview.Auth;

namespace SiteReviewProB
{
    public static class SiteReviewProB
    {
        [FunctionName("SiteReviewProB")]
        public static async Task<IActionResult> RunProB(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req,
        ILogger log, ExecutionContext executionContext)
        {
            log.LogInformation($"SiteReviewProB timer trigger function executed at: {DateTime.Now}");

            try
            {
                var config = new ConfigurationBuilder()
                    .SetBasePath(executionContext.FunctionAppDirectory)
                    .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                    .AddEnvironmentVariables()
                    .Build();

                var scopes = new[] { "user.read", "mail.send" };
                ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
                var graphClient = new GraphServiceClient(auth, scopes);

                log.LogInformation("Graph API authentication successful.");

                var sites = await GetProtectedBSites(graphClient, log);
                log.LogInformation($"Retrieved {sites.Count} protected B sites.");

                var publicSites = new List<Site>();

                foreach (var site in sites)
                {
                    log.LogInformation($"Processing site: {site.DisplayName}");

                    var sitePrivacySetting = await GetSitePrivacySetting(graphClient, site, log);
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

                    var emailBody = GenerateEmailBodyProtectedB(publicSites);

                    await SendEmailProB(recipientEmails, emailBody, graphClient, log);
                }

                log.LogInformation("Function app executed successfully");
            }
            catch (Exception ex)
            {
                log.LogError($"An error occurred: {ex.Message}");
            }

            return new OkObjectResult("HTTP trigger executed successfully.");
        }

        private static async Task<List<Site>> GetProtectedBSites(GraphServiceClient graphClient, ILogger log)
        {
            var sites = new List<Site>();
            log.LogInformation("Getting Protected B Sites.");

            try
            {
                var siteCollectionPage = await graphClient.Sites.Request().GetAsync();
                while (siteCollectionPage != null)
                {
                    log.LogInformation($"Processing {siteCollectionPage.Count} sites from current page.");
                    foreach (var site in siteCollectionPage)
                    {
                        log.LogInformation($"Checking site: {site.DisplayName}, URL: {site.WebUrl}");
                        if (site.WebUrl.Contains("/teams/b"))
                        {
                            log.LogInformation($"Site {site.DisplayName} identified as Protected B.");
                            sites.Add(site);
                        }
                    }
                    if (siteCollectionPage.NextPageRequest != null)
                    {
                        siteCollectionPage = await siteCollectionPage.NextPageRequest.GetAsync();
                    }
                    else
                    {
                        siteCollectionPage = null;
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

        private static string GenerateEmailBodyProtectedB(List<Site> publicSites)
        {
            var emailBody = "The following Protected B sites are set to public:<br>";
            foreach (var site in publicSites)
            {
                emailBody += $"{site.DisplayName} - {site.WebUrl}<br>";
            }
            return emailBody;
        }

        private static async Task SendEmailProB(string[] recipientEmails, string emailBody, GraphServiceClient graphClient, ILogger log)
        {
            var message = new Message
            {
                Subject = "Protected B Sites Report",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = emailBody
                },
                ToRecipients = recipientEmails.Select(email => new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = email
                    }
                }).ToList()
            };

            await graphClient.Me.SendMail(message, false).Request().PostAsync();
            log.LogInformation("Report email sent successfully.");
        }
    }
}