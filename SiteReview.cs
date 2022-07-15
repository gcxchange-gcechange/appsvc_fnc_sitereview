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
            
            // Get a report of site usage in the last 180 days
            var reportMsg = graphAPIAuth.Reports
            .GetSharePointSiteUsageDetail("D180")
            .Request()
            .Header("ConsistencyLevel", "eventual")
            .GetHttpRequestMessage();

            // Download the CSV data
            var reportResponse = await graphAPIAuth.HttpProvider.SendAsync(reportMsg);
            var reportCSV = await reportResponse.Content.ReadAsStringAsync();

            var report = Helpers.GenerateCSV(reportCSV); 

            // Look at the header for the index of data we care about
            var siteIdIndex = report[0].FindIndex(l => l.Equals("Site Id"));
            var lastActivityIndex = report[0].FindIndex(l => l.Equals("Last Activity Date"));

            // Skip over any excluded sites
            var excludeSiteIds = Globals.GetExcludedSiteIds();

            var warningSiteIds = new List<string>();
            var deleteSiteIds = new List<string>();

            // Build the list of warning and delete sites
            for (var i = 1; i < report.Count; i++)
            {
                var lastActivityDate = report[i][lastActivityIndex];

                if (lastActivityDate != string.Empty)
                {
                    if (excludeSiteIds.Contains(report[i][siteIdIndex]))
                        continue;

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

            // Send warning emails to site owners
            foreach (var id in warningSiteIds)
            {
                var site = graphAPIAuth.Sites[id]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;

                if (site != null)
                {
                    var groupQueryOptions = new List<QueryOption>()
                    {
                        new QueryOption("$search", "\"mailNickname:" + site.Name +"\"")
                    };

                    var groups = await graphAPIAuth.Groups
                    .Request(groupQueryOptions)
                    .Header("ConsistencyLevel", "eventual")
                    .GetAsync();

                    do
                    {
                        foreach (var group in groups)
                        {
                            var owners = await graphAPIAuth.Groups[group.Id].Owners
                            .Request()
                            .GetAsync();

                            do
                            {
                                foreach (var owner in owners)
                                {
                                    var user = await graphAPIAuth.Users[owner.Id]
                                    .Request()
                                    .Select("displayName,mail")
                                    .GetAsync();

                                    if (user != null)
                                    {
                                        await Email.SendWarningEmail(user.Mail, log);
                                    }
                                }
                            }
                            while (owners.NextPageRequest != null && (owners = await owners.NextPageRequest.GetAsync()).Count > 0);
                        }
                    }
                    while (groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0);
                }
            }

            // Delete sites and inform owners
            foreach (var id in deleteSiteIds)
            {
                var appOnlyAuth = auth.appOnlyAuth("https://devgcx.sharepoint.com/", log);
                
            }














            // Use this when it's out of beta
            // https://docs.microsoft.com/en-us/graph/api/teams-list?view=graph-rest-beta&tabs=csharp

            var teamsGroups = await graphAPIAuth.Groups
            .Request()
            .Header("ConsistencyLevel", "eventual")
            .Filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
            .GetAsync();

            do
            {
                foreach (var group in teamsGroups)
                {
                    // group id is the same as team id

                    //https://docs.microsoft.com/en-us/graph/api/group-delete?view=graph-rest-1.0&tabs=csharp

                    var owners = await graphAPIAuth.Groups[group.Id].Owners
                    .Request()
                    .Header("ConsistencyLevel", "eventual")
                    .GetAsync();
                }
            }
            while (teamsGroups.NextPageRequest != null && (teamsGroups = await teamsGroups.NextPageRequest.GetAsync()).Count > 0);

            return new OkObjectResult("Function app executed successfully");
        }
    }
}
