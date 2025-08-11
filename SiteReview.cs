using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Azure.Functions.Worker;
using static SiteReview.Auth;

namespace SiteReview
{
    public static class SiteReview
    {
        [Function("SiteReview")]
        public static async Task Run(
            [TimerTrigger("0 0 0 * * 6")] TimerInfo myTimer, FunctionContext executionContext)
        {
            var log = executionContext.GetLogger("SiteReview");
            log.LogInformation($"SiteReview executed at {DateTime.Now} with report only mode: {Globals.reportOnlyMode}");

            var graphAPIAuth = new Auth().graphAuth(log);
            var report = await Common.GetReport(graphAPIAuth, log);

            log.LogInformation($"Found {report.NoOwnerSites.Count} sites with less than {Globals.minSiteOwners} owners.");
            log.LogInformation($"Found {report.StorageThresholdSites.Count} sites over {Globals.storageThreshold}% storage capacity.");
            log.LogInformation($"Found {report.WarningSites.Count} sites inactive for {Globals.inactiveDaysWarn} days but less than {Globals.inactiveDaysDelete} days.");
            log.LogInformation($"Found {report.DeleteSites.Count} sites inactive for {Globals.inactiveDaysDelete} days or more.");
            log.LogInformation($"Found {report.PrivacySettingSites.Count} sites that we not set to private.");
            log.LogInformation($"Found {report.ClassificationSites.Count} sites that had no classification.");
            log.LogInformation($"Found {report.HubAssociationSites.Count} sites that were not associated with the hub site {Globals.hubId}.");

            var scopes = new[] { "user.read mail.send" };
            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
            var graphClient = new GraphServiceClient(auth, scopes);

            await Email.SendReportEmail(Globals.adminEmails, report, graphClient, log);

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
                    var s = await graphAPIAuth
                    .Sites[site.SiteId]
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.Headers.Add("ConsistencyLevel", "eventual");
                    });

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

                await StoreData.StoreSitesToDelete(deleteSiteIds, Common.DeleteSiteIdsContainerName, log);
            }

            //return new OkObjectResult("Function app executed successfully");
            log.LogInformation("Function app executed successfully");
        }
    }
}
