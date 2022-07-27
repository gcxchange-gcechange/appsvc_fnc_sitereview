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
    internal class DeleteSite
    {
        [FunctionName("DeleteSites")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation($"SiteReview executed at {DateTime.Now}");

            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var report = await Common.GetReport(graphAPIAuth, log); 

            foreach (var site in report.DeleteSites)
            {
                var s = graphAPIAuth.Sites[site.SiteId]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;

                if (s != null)
                {
                    await Common.DeleteSite(site.SiteUrl, graphAPIAuth, log);
                }
            }

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
