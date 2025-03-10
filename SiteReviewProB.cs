using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using SiteReview;

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
                var graphAPIAuth = new Auth().graphAuth(log);
                var sites = await GetProtectedBSites(graphAPIAuth, log);

                var publicSites = new List<Site>();

                foreach (var site in sites)
                {
                    var sitePrivacySetting = await GetSitePrivacySetting(graphAPIAuth, site.Id, log);
                    if (sitePrivacySetting == "Public")
                    {
                        publicSites.Add(site);
                    }
                }

                if (publicSites.Any())
                {
                    await SendReportEmail(publicSites, graphAPIAuth, log);
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

            try
            {
                var siteCollectionPage = await graphClient.Sites.Request().GetAsync();
                while (siteCollectionPage != null)
                {
                    sites.AddRange(siteCollectionPage.Where(site => site.WebUrl.Contains("/teams/b")));
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

            return sites;
        }

        private static async Task<string> GetSitePrivacySetting(GraphServiceClient graphClient, string siteId, ILogger log)
        {
            try
            {
                var site = await graphClient.Sites[siteId].Request().GetAsync();
                var siteDrive = await graphClient.Sites[siteId].Drive.Request().GetAsync();
                var permissions = await graphClient.Sites[siteId].Permissions.Request().GetAsync();
                var isPublic = permissions.Any(p => p.Roles.Contains("read") && p.GrantedToV2.User != null);
                return isPublic ? "Public" : "Private";
            }
            catch (Exception ex)
            {
                log.LogError($"Failed to get privacy setting for site {siteId}: {ex.Message}");
                return "Unknown";
            }
        }

        private static async Task SendReportEmail(List<Site> publicSites, GraphServiceClient graphClient, ILogger log)
        {
            var emailAddresses = new List<string> { "admin1@example.com", "admin2@example.com" }; // need to replace with actual email addresses

            var emailBody = "The following Protected B sites are set to public:\n\n";
            foreach (var site in publicSites)
            {
                emailBody += $"- {site.DisplayName} ({site.WebUrl})\n";
            }

            foreach (var email in emailAddresses)
            {
                var message = new Message
                {
                    Subject = "Public Protected B Sites Report",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = emailBody
                    },
                    ToRecipients = new List<Recipient>
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = email
                            }
                        }
                    }
                };

                try
                {
                    await graphClient.Me.SendMail(message).Request().PostAsync();
                    log.LogInformation($"Report email sent to {email}");
                }
                catch (Exception ex)
                {
                    log.LogError($"Failed to send report email to {email}: {ex.Message}");
                }
            }
        }
    }
}