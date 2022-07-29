using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SiteReview
{
    public static class Common
    {
        public static async Task<SiteReport> GetReport(GraphServiceClient graphAPIAuth, ILogger log)
        {
            // Get a report of site usage in the last 180 days
            var reportMsg = graphAPIAuth.Reports
            .GetSharePointSiteUsageDetail("D180")
            .Request()
            .Header("ConsistencyLevel", "eventual")
            .GetHttpRequestMessage();

            log.LogInformation("Got site usage report.");

            // Download the CSV data
            var reportResponse = await graphAPIAuth.HttpProvider.SendAsync(reportMsg);
            var reportCSV = await reportResponse.Content.ReadAsStringAsync();

            var report = Helpers.GenerateCSV(reportCSV);

            // Look at the header for the index of data we care about
            var siteIdIndex = report.FirstOrDefault().FindIndex(l => l.Equals("Site Id"));
            var lastActivityIndex = report.FirstOrDefault().FindIndex(l => l.Equals("Last Activity Date"));
            var siteURLIndex = report.FirstOrDefault().FindIndex(l => l.Equals("Site URL"));

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

            var siteReport = new SiteReport();
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
                        var reportData = new ReportData(siteId, siteURL);
                        var daysInactive = (DateTime.Now - DateTime.Parse(lastActivityDate)).TotalDays;

                        if (daysInactive > 120)
                        {
                            siteReport.DeleteSites.Add(reportData);
                            log.LogWarning($"Flagged for deletion: {siteURL}");
                        }
                        else if (daysInactive > 60)
                        {
                            siteReport.WarningSites.Add(reportData);
                            log.LogWarning($"Flagged for warning: {siteURL}");
                        }
                    }
                }
            }

            return siteReport;
        }

            public static async Task<List<User>> GetSiteOwners(string siteId, GraphServiceClient graphAPIAuth)
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

        public static async Task<bool> DeleteSiteGroup(string siteUrl, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var success = true;

            try
            {
                var ctx = new Auth().appOnlyAuth(siteUrl, log);
                ctx.Load(ctx.Site, s => s.GroupId);
                ctx.ExecuteQuery();

                var groupId = ctx.Site.GroupId;

                await graphAPIAuth.Groups[groupId.ToString()]
                .Request()
                .DeleteAsync();
            }
            catch (Exception ex)
            {
                log.LogError($"Error finding and deleting m365 group for {siteUrl} - {ex}");
                success = false;
            }

            return success;
        }

        public static async Task<bool> DeleteSite(string siteUrl, ILogger log)
        {
            var success = true;

            try
            {
                var ctx = new Auth().appOnlyAuth("https://devgcx-admin.sharepoint.com/", log);
                var tenant = new Tenant(ctx);
                var removeSite = tenant.RemoveSite(siteUrl);
                ctx.Load(removeSite);
                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.LogError($"Error deleting {siteUrl} - {ex}");
                success = false;
            }

            return success;
        }

        public class SiteReport
        {
            public List<ReportData> WarningSites { get; set; }
            public List<ReportData> DeleteSites { get; set; }
        }

        public class ReportData
        {
            public ReportData(string siteId, string siteUrl)
            {
                SiteId = siteId;
                SiteUrl = siteUrl;
            }

            public string SiteId { get; set; }
            public string SiteUrl { get; set; }
        }

        //var teamsGroups = await graphAPIAuth.Groups
        //.Request()
        //.Header("ConsistencyLevel", "eventual")
        //.Filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
        //.GetAsync();
        //
        //do
        //{
        //    foreach (var group in teamsGroups)
        //    {
        //        // group id is the same as team id
        //
        //        //https://docs.microsoft.com/en-us/graph/api/group-delete?view=graph-rest-1.0&tabs=csharp
        //
        //        var owners = await graphAPIAuth.Groups[group.Id].Owners
        //        .Request()
        //        .Header("ConsistencyLevel", "eventual")
        //        .GetAsync();
        //    }
        //}
        //while (teamsGroups.NextPageRequest != null && (teamsGroups = await teamsGroups.NextPageRequest.GetAsync()).Count > 0);
    }
}
