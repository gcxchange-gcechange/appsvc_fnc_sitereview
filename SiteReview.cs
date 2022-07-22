using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Online.SharePoint.TenantAdministration;
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
            log.LogInformation($"SiteReview executed at {DateTime.Now}");

            var auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            
            // Get a report of site usage in the last 180 days
            var reportMsg = graphAPIAuth.Reports
            .GetSharePointSiteUsageDetail("D180") //GetTeamsUserActivityUserDetail - GetSharePointSiteUsageDetail
            .Request()
            .Header("ConsistencyLevel", "eventual")
            .GetHttpRequestMessage();

            log.LogInformation("Got site usage report.");

            // Download the CSV data
            var reportResponse = await graphAPIAuth.HttpProvider.SendAsync(reportMsg);
            var reportCSV = await reportResponse.Content.ReadAsStringAsync();

            var report = Helpers.GenerateCSV(reportCSV);

            // Look at the header for the index of data we care about
            var siteIdIndex = report[0].FindIndex(l => l.Equals("Site Id"));
            var lastActivityIndex = report[0].FindIndex(l => l.Equals("Last Activity Date"));
            var siteURLIndex = report[0].FindIndex(l => l.Equals("Site URL"));

            var excludeSiteIds = Globals.GetExcludedSiteIds();
            var subsiteIds = new List<string>();

            // Get subsites
            var sitesQueryOptions = new List<QueryOption>()
            {
                new QueryOption("search", "DepartmentId:{" + Globals.hubId + "}"),
            };

            var subsites = await graphAPIAuth.Sites
            .Request(sitesQueryOptions)
            .Header("ConsistencyLevel", "eventual")
            .GetAsync();

            do
            {
                foreach (var site in subsites)
                {
                    subsiteIds.Add(site.Id.Split(",")[1]);
                }
            }
            while (subsites.NextPageRequest != null && (subsites = await subsites.NextPageRequest.GetAsync()).Count > 0);

            var warningSites = new List<Tuple<string, string>>();
            var deleteSites = new List<Tuple<string, string>>();

            // Build the list of warning and delete sites
            for (var i = 1; i < report.Count; i++)
            {
                var siteId = report[i][siteIdIndex];
                var lastActivityDate = report[i][lastActivityIndex];
                var siteURL = report[i][siteURLIndex];

                if (lastActivityDate != string.Empty)
                {
                    // Skip excluded sites
                    if (excludeSiteIds.Contains(siteId))
                        continue;

                    // If the site is a subsite
                    if (subsiteIds.Any(s => s == siteId))
                    {
                        var siteData = new Tuple<string, string>(siteId, siteURL);
                        var daysInactive = (DateTime.Now - DateTime.Parse(lastActivityDate)).TotalDays;

                        if (daysInactive > 120)
                        {
                            deleteSites.Add(siteData);
                            log.LogWarning($"Flagged for deletion: {siteURL}");
                        }
                        else if (daysInactive > 60)
                        {
                            warningSites.Add(siteData);
                            log.LogWarning($"Flagged for warning: {siteURL}");
                        }
                    }
                }
            }

            log.LogInformation($"Discovered {warningSites.Count} sites flagged for warning.");
            log.LogInformation($"Discovered {deleteSites.Count} sites flagged for deletion.");

            // Send warning emails to site owners
            foreach (var site in warningSites)
            {
                var siteOwners = await GetSiteOwners(site.Item1, graphAPIAuth);

                if (siteOwners.Count > 0)
                {
                    foreach (var owner in siteOwners)
                    {
                        await Email.SendWarningEmail(owner.Mail, site.Item2, log);
                    }
                }
                else
                {
                    await Email.SendWarningEmail("gcxgce-admin", site.Item2, log); // TODO: Get correct email address
                }
            }

            // Delete sites and inform owners
            foreach (var site in deleteSites)
            {
                var siteOwners = await GetSiteOwners(site.Item1, graphAPIAuth);

                if (siteOwners.Count > 0)
                {
                    foreach (var owner in siteOwners)
                    {
                        await Email.SendDeleteEmail(owner.Mail, site.Item2, log);
                    }
                }
                else
                {
                    await Email.SendDeleteEmail("gcxgce-admin", site.Item2, log); // TODO: Get correct email address
                }

                var s = graphAPIAuth.Sites[site.Item1]
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync()
                .Result;

                if (s != null)
                {
                    //var ctx = auth.appOnlyAuth("https://devgcx.sharepoint.com/", log);
                    //var tenant = new Tenant(ctx);
                    //
                    //var removeSite = tenant.RemoveSite(s.WebUrl);
                    //ctx.Load(removeSite);
                    //ctx.ExecuteQuery();
                }

            }






            // https://www.codesharepoint.com/csom/delete-sub-site-in-sharepoint-using-csom







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

        private static async Task<List<User>> GetSiteOwners(string siteId, GraphServiceClient graphAPIAuth)
        {
            var siteOwners = new List<User>();

            var site = graphAPIAuth.Sites[siteId]
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
                                    siteOwners.Add(user);
                                }
                            }
                        }
                        while (owners.NextPageRequest != null && (owners = await owners.NextPageRequest.GetAsync()).Count > 0);
                    }
                }
                while (groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0);
            }

            return siteOwners;
        }
    }
}
