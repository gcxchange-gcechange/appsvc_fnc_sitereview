using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace SiteReview
{
    internal class DeleteSite
    {
        [FunctionName("DeleteSites")]
        public static async Task<IActionResult> Run(
            [TimerTrigger("0 0 0 2 1-12 *")] TimerInfo myTimer, ILogger log, ExecutionContext executionContext)
        {
            log.LogInformation($"DeleteSites executed at {DateTime.Now}");

            var graphAPIAuth = new Auth().graphAuth(log);

            var siteIds = await StoreData.GetSitesToDelete(executionContext, Common.DeleteSiteIdsContainerName, log);

            log.LogInformation($"Found {siteIds.Count} sites to be deleted");

            foreach (var id in siteIds)
            {
                var site = graphAPIAuth.Sites[id]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;
                
                if (site != null)
                {
                    await Common.DeleteSite(site.WebUrl, log);
                }
            }

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
