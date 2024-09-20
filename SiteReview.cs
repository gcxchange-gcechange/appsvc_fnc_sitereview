using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Extensions.Http;

namespace SiteReview
{
    public static class SiteReview
    {
        [FunctionName("SiteReview")]
        public static async Task<IActionResult> Run(
            //[TimerTrigger("0 0 0 1 1-12 *")] TimerInfo myTimer, ILogger log, ExecutionContext executionContext)
            [HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req, ILogger log, ExecutionContext executionContext)
        {
            log.LogInformation($"SiteReview executed at {DateTime.Now} with report only mode: {Globals.reportOnlyMode}");

            var graphAPIAuth = new Auth().graphAuth(log);

            var report = await Common.GetReport(graphAPIAuth, log);

            log.LogInformation($"Found {report.NoOwnerSites.Count} sites with less than {Globals.minSiteOwners} owners.");
            log.LogInformation($"Found {report.StorageThresholdSites.Count} sites over {Globals.storageThreshold}% storage capacity.");
            log.LogInformation($"Found {report.WarningSites.Count + report.DeleteSites.Count} sites inactive for {Globals.inactiveDaysWarn} days or more.");
            log.LogInformation($"Found {report.DeleteSites.Count} sites inactive for {Globals.inactiveDaysDelete} days or more.");
            log.LogInformation($"Found {report.PrivacySettingSites.Count} sites there we not set to private.");
            log.LogInformation($"Found {report.ClassificationSites.Count} sites there had no classification.");

            var combinedReportSites = report.WarningSites
                .Concat(report.DeleteSites)
                .Concat(report.NoOwnerSites)
                .Concat(report.StorageThresholdSites)
                .Concat(report.PrivacySettingSites)
                .Concat(report.ClassificationSites)
                .GroupBy(site => site.SiteId)
                .Select(site => site.First())
                .ToList();

            // Send the report to the admin email address
            await Email.SendReportEmail(Globals.adminEmails, combinedReportSites, graphAPIAuth, log);

            if (!Globals.reportOnlyMode)
            {
                // Send warning emails to site owners
                foreach (var site in report.WarningSites)
                {
                    if (site.SiteOwners.Count > 0)
                    {
                        foreach (var owner in site.SiteOwners)
                        {
                            await Email.SendWarningEmail(owner.Mail, site.SiteUrl, graphAPIAuth, log);
                        }
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
                            if (site.SiteOwners.Count > 0)
                            {
                                foreach (var owner in site.SiteOwners)
                                {
                                    await Email.SendDeleteEmail(owner.Mail, site.SiteUrl, graphAPIAuth, log);
                                }
                            }

                            deleteSiteIds.Add(site.SiteId);
                        }
                    }
                }

                await StoreData.StoreSitesToDelete(executionContext, deleteSiteIds, Common.DeleteSiteIdsContainerName, log);
            }

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
