using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;

namespace SiteReview
{
    public static class SiteReview
    {
        [FunctionName("InformOwnersAndDeleteGroups")]
        public static async Task<IActionResult> Run(
            [TimerTrigger("0 0 0 1 1-12 *")] TimerInfo myTimer, ILogger log, ExecutionContext executionContext)
        {
            log.LogInformation($"InformOwnersAndDeleteGroups executed at {DateTime.Now}");

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

            var deleteSiteIds = new List<string>();

            // Delete groups and inform owners
            foreach (var site in report.DeleteSites)
            {
                var s = graphAPIAuth.Sites[site.SiteId]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;
            
                if (s != null)
                {
                    var deleteSuccess = await Common.DeleteSiteGroup(site.SiteUrl, graphAPIAuth, log);

                    if (deleteSuccess)
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
            
                        deleteSiteIds.Add(site.SiteId);
                    }
                }
            }

            await StoreData.StoreSitesToDelete(executionContext, deleteSiteIds, Common.DeleteSiteIdsContainerName, log);

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
