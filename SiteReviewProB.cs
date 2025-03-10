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
                var group = await Common.GetGroupFromSite(site, graphClient, log);

                if (group != null)
                {
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
            var userEmails = new[] { "email1@example.com", "email2@example.com" }; // add the email addresses to be emailed to

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
        }

        private static async Task<bool> SendEmailWrapper(string userEmail, string emailSubject, string emailBody, BodyType bodyType, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var emailType = typeof(Email);
            var sendEmailMethod = emailType.GetMethod("SendEmail", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            if (sendEmailMethod != null)
            {
                var task = (Task<bool>)sendEmailMethod.Invoke(null, new object[] { userEmail, emailSubject, emailBody, bodyType, graphAPIAuth, log });
                return await task;
            }
            log.LogError("Failed to invoke SendEmail method.");
            return false;
        }

    }
}