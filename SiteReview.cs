using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.Graph;
using System.Net;
using System.Linq;

namespace SiteReview
{
    public static class SiteReview
    {
        [FunctionName("SiteReview")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var reportMsg = graphAPIAuth.Reports
            .GetSharePointSiteUsageDetail("D180")
            .Request()
            .Header("ConsistencyLevel", "eventual")
            .GetHttpRequestMessage();

            var reportResponse = await graphAPIAuth.HttpProvider.SendAsync(reportMsg);
            var reportCSV = await reportResponse.Content.ReadAsStringAsync();

            var report = Helpers.GenerateCSV(reportCSV); 

            var siteIdIndex = report[0].FindIndex(l => l.Equals("Site Id"));
            var lastActivityIndex = report[0].FindIndex(l => l.Equals("Last Activity Date"));

            var warningSiteIds = new List<string>();
            var deleteSiteIds = new List<string>();

            for (var i = 1; i < report.Count; i++)
            {
                var lastActivityDate = report[i][lastActivityIndex];

                if (lastActivityDate != string.Empty)
                {
                    var daysInactive = (DateTime.Now - DateTime.Parse(lastActivityDate)).TotalDays;

                    if (daysInactive > 120)
                    {
                        deleteSiteIds.Add(report[i][siteIdIndex]);
                    }
                    else if (daysInactive > 60)
                    {
                        warningSiteIds.Add(report[i][siteIdIndex]);
                    }
                }
            }

            var excludeSiteIds = Globals.GetExcludedSiteIds();

            foreach (var id in warningSiteIds)
            {
                if (excludeSiteIds.Contains(id))
                    continue;

                var site = graphAPIAuth.Sites[id]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;
            }

            foreach (var id in deleteSiteIds)
            {
                if (excludeSiteIds.Contains(id))
                    continue;

                var appOnlyAuth = auth.appOnlyAuth("https://devgcx.sharepoint.com/", log);
                
            }

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
