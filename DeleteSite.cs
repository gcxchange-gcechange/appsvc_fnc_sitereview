using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace SiteReview
{
    internal class DeleteSite
    {
        [Function("DeleteSites")]
        public static async Task Run(
            [TimerTrigger("0 0 0 2 1-12 *")] TimerInfo myTimer, FunctionContext executionContext)
        {
            var log = executionContext.GetLogger("DeleteSites");
            log.LogInformation($"DeleteSites executed at {System.DateTime.Now}");

            if (!Globals.reportOnlyMode)
            {
                var graphAPIAuth = new Auth().graphAuth(log);

                var siteIds = await StoreData.GetSitesToDelete(Common.DeleteSiteIdsContainerName, log);

                log.LogInformation($"Found {siteIds.Count} sites to be deleted");

                foreach (var id in siteIds)
                {
                    var site = await graphAPIAuth
                    .Sites[id]
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.Headers.Add("ConsistencyLevel", "eventual");
                    });

                    if (site != null)
                    {
                        Common.DeleteSite(site.WebUrl, log);
                    }
                }
            }

            log.LogInformation("Function app executed successfully");
        }
    }
}
