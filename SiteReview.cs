using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace SiteReview
{
    public static class SiteReview
    {
        [FunctionName("InformOwnersAndDeleteGroups")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation($"SiteReview executed at {DateTime.Now}");

            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var report = await Common.GetReport(graphAPIAuth, log);

            log.LogInformation($"Discovered {report.WarningSites.Count} sites flagged for warning.");
            log.LogInformation($"Discovered {report.DeleteSites.Count} sites flagged for deletion.");

            // Send warning emails to site owners
            foreach (var site in report.WarningSites)
            {
                var siteOwners = await Common.GetSiteOwners(site.SiteId, graphAPIAuth);

                if (siteOwners.Count > 0)
                {
                    foreach (var owner in siteOwners)
                    {
                        await Email.SendWarningEmail(owner.Mail, site.SiteUrl, log);
                    }
                }
                else
                {
                    await Email.SendWarningEmail("gcxgce-admin@tbs-sct.gc.ca", site.SiteUrl, log);
                }
            }

            // Delete sites and inform owners
            foreach (var site in report.DeleteSites)
            {
                var siteOwners = await Common.GetSiteOwners(site.SiteId, graphAPIAuth);

                if (siteOwners.Count > 0)
                {
                    foreach (var owner in siteOwners)
                    {
                        await Email.SendDeleteEmail(owner.Mail, site.SiteUrl, log);
                    }
                }
                else
                {
                    await Email.SendDeleteEmail("gcxgce-admin@tbs-sct.gc.ca", site.SiteUrl, log);
                }

                var s = graphAPIAuth.Sites[site.SiteId]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;

                if (s != null)
                {
                    await Common.DeleteSiteGroup(site.SiteUrl, graphAPIAuth, log);
                }
            }

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
