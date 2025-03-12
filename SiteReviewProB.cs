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
                log.LogInformation("Graph API authentication successful.");

                var sites = await GetProtectedBSites(graphAPIAuth, log);
                log.LogInformation($"Retrieved {sites.Count} protected B sites.");

                var publicSites = new List<Site>();

                foreach (var site in sites)
                {
                    var sitePrivacySetting = await GetSitePrivacySetting(graphAPIAuth, site.Id, log);
                    log.LogInformation($"Site {site.Id} privacy setting: {sitePrivacySetting}");

                    if (sitePrivacySetting == "Public")
                    {
                        publicSites.Add(site);
                        log.LogInformation($"Site {site.Id} added to public sites list.");
                    }
                }

                if (publicSites.Any())
                {
                    log.LogInformation("Public sites found, sending report email.");
                    await SendReportEmailProB(publicSites, graphAPIAuth, log);
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

            log.LogInformation($"Total Protected B Sites retrieved: {sites.Count}");
            return sites;
        }

        private static async Task<string> GetSitePrivacySetting(GraphServiceClient graphClient, string siteId, ILogger log)
        {
            log.LogInformation($"Getting privacy setting for site: {siteId}");
            try
            {
                var site = await graphClient.Sites[siteId].Request().GetAsync();
                var group = await Common.GetGroupFromSite(site, graphClient, log);

                if (group != null)
                {
                    log.LogInformation($"Group found for site: {siteId}, privacy setting: {group.Visibility}");
                    return group.Visibility;
                }
                else
                {
                    log.LogWarning($"Group not found for siteId {siteId}");
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Failed to get site privacy setting for siteId {siteId}: {ex.Message}");
            }

            return "Unknown";
        }

        private static async Task SendReportEmailProB(List<Site> publicSites, GraphServiceClient graphAPIAuth, ILogger log)
        {
            log.LogInformation("Preparing to send report email for public sites.");
            var userEmails = new[] { "email1@example.com", "email2@example.com" }; // need add the email addresses to be emailed to

            var siteDetails = publicSites.Select(site =>
                $"Site Name: {site.DisplayName}<br>Site URL: <a href='{site.WebUrl}'>{site.WebUrl}</a><br><br>"
            );

            var emailBody = $@"
                Greetings,<br><br>
                The following Protected B sites have their privacy setting set to public:<br><br>
                {string.Join("<hr>", siteDetails)}
                Please review these sites and take necessary actions.<br><br>
                Regards,<br>The GCX Team";

            List<Task> emailTasks = new List<Task>();

            foreach (var email in userEmails)
            {
                log.LogInformation($"Sending email to: {email}");
                emailTasks.Add(SendEmailWrapper(
                    email,
                    "Protected B Sites Public Privacy Setting Report",
                    emailBody,
                    BodyType.Html,
                    graphAPIAuth,
                    log
                ));
            }

            await Task.WhenAll(emailTasks);
            log.LogInformation("Report email sent successfully.");
        }

        private static async Task<bool> SendEmailWrapper(string userEmail, string emailSubject, string emailBody, BodyType bodyType, GraphServiceClient graphAPIAuth, ILogger log)
        {
            log.LogInformation($"Invoking SendEmail method for user: {userEmail}");
            var emailType = typeof(Email);
            var sendEmailMethod = emailType.GetMethod("SendEmail", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

            if (sendEmailMethod != null)
            {
                var task = (Task<bool>)sendEmailMethod.Invoke(null, new object[] { userEmail, emailSubject, emailBody, bodyType, graphAPIAuth, log });
                log.LogInformation($"Email send task for user: {userEmail} started.");
                return await task;
            }
            log.LogError("Failed to invoke SendEmail method.");
            return false;
        }
    }
}