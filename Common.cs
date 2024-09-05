using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace SiteReview
{
    public static class Common
    {
        public static readonly string DeleteSiteIdsContainerName = "delete";
        public static async Task<SiteReport> GetReport(GraphServiceClient graphAPIAuth, ILogger log)
        {
            var siteReport = new SiteReport();

            try
            {
                // Get a report of site usage in the last 180 days
                var siteReportMsg = graphAPIAuth.Reports
                .GetSharePointSiteUsageDetail("D180")
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetHttpRequestMessage();

                log.LogInformation("Got site usage report.");

                // Download the site CSV data
                var siteReportResponse = await graphAPIAuth.HttpProvider.SendAsync(siteReportMsg);
                var siteCSV = Helpers.GenerateCSV(await siteReportResponse.Content.ReadAsStringAsync());

                // Look at the site CSV header for the index of data we care about
                var siteSiteIdIndex = siteCSV.FirstOrDefault().FindIndex(l => l.Equals("Site Id"));
                var siteLastActivityIndex = siteCSV.FirstOrDefault().FindIndex(l => l.Equals("Last Activity Date"));
                var siteSiteURLIndex = siteCSV.FirstOrDefault().FindIndex(l => l.Equals("Site URL"));
                var siteStorageUsedIndex = siteCSV.FirstOrDefault().FindIndex(l => l.Equals("Storage Used (Byte)"));
                var siteStorageAllocatedIndex = siteCSV.FirstOrDefault().FindIndex(l => l.Equals("Storage Allocated (Byte)"));

                // Get the teams usage report
                var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityDetail(period='D180')");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await new Auth().GetAccessTokenAsync());

                var httpClient = new HttpClient();
                var response = await httpClient.SendAsync(request);

                var teamsActivityCSV = new List<List<string>>();
                if (response.IsSuccessStatusCode)
                {
                    log.LogInformation("Got teams usage report.");
                    teamsActivityCSV = Helpers.GenerateCSV(await response.Content.ReadAsStringAsync());
                }
                else
                {
                    log.LogError($"Error retrieving teams usage report: {response.StatusCode}");
                }

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

                // Build the report
                for (var i = 1; i < siteCSV.Count; i++)
                {
                    var siteId = siteCSV[i][siteSiteIdIndex];
                    var lastActivityDate = siteCSV[i][siteLastActivityIndex];
                    var siteURL = siteCSV[i][siteSiteURLIndex];
                    var storageUsed = siteCSV[i][siteStorageUsedIndex];
                    var storageAllocated = siteCSV[i][siteStorageAllocatedIndex];

                    // TODO: Do we care if there's no last activity date??
                    if (lastActivityDate != string.Empty)
                    {
                        // Skip excluded sites
                        if (excludeSiteIds.Contains(siteId))
                            continue;

                        // If the site is a subsite
                        if (subsiteIds.Any(s => s == siteId))
                        {
                            var siteOwners = await GetSiteOwners(siteId, graphAPIAuth, log);

                            var site = await graphAPIAuth.Sites[siteId].Request().GetAsync();
                            siteURL = site != null ? site.WebUrl : $"Error: The WebUrl for SiteId {siteId} could not be found.";

                            var siteDaysInactive = (DateTime.Now - DateTime.Parse(lastActivityDate)).TotalDays;
                            var teamDaysInactive = GetTeamsActivity(teamsActivityCSV, site.DisplayName, log);

                            var reportData = new ReportData(
                                siteId, 
                                siteURL, 
                                (int)Math.Min(siteDaysInactive, teamDaysInactive), 
                                siteOwners, 
                                ulong.Parse(storageAllocated), 
                                ulong.Parse(storageUsed)
                            );

                            // If site and teams activity meets our threshold, add to the report.
                            if (siteDaysInactive >= Globals.inactiveDaysDelete && teamDaysInactive >= Globals.inactiveDaysDelete)
                            {
                                siteReport.DeleteSites.Add(reportData);
                            }
                            else if (siteDaysInactive >= Globals.inactiveDaysWarn && teamDaysInactive >= Globals.inactiveDaysWarn)
                            {
                                siteReport.WarningSites.Add(reportData);
                            }

                            if (siteOwners.Count < Globals.minSiteOwners)
                                siteReport.NoOwnerSites.Add(reportData);

                            var usedPercentage = double.Parse(storageUsed) / double.Parse(storageAllocated) * 100;
                            if (usedPercentage >= Globals.storageThreshold)
                                siteReport.StorageThresholdSites.Add(reportData);
                        }
                    }
                }

                return siteReport;
            }
            catch (Exception ex)
            {
                log.LogError($"Error building report - {ex.Message} - {ex.StackTrace}");
                return siteReport;
            }
        }

        private static async Task<List<User>> GetSiteOwners(string siteId, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var siteOwners = new List<User>();

            try
            {
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
            catch (Exception ex)
            {
                log.LogError($"Error getting site owners for siteId {siteId} - {ex.Message}");
                return siteOwners;
            }
        }

        private static double GetTeamsActivity(List<List<string>> teamsActivityCSV, string siteDisplayName, ILogger log)
        {
            try
            {
                var teamNameIndex = teamsActivityCSV.FirstOrDefault().FindIndex(l => l.Equals("Team Name"));
                var lastActivityIndex = teamsActivityCSV.FirstOrDefault().FindIndex(l => l.Equals("Last Activity Date"));

                for (var i = 1; i < teamsActivityCSV.Count; i++)
                {
                    // Remove this
                    Console.WriteLine(teamsActivityCSV[i][teamNameIndex]);

                    if (teamsActivityCSV[i][teamNameIndex] == siteDisplayName)
                    {
                        var teamLastActivityDate = teamsActivityCSV[i][lastActivityIndex];
                        return (DateTime.Now - DateTime.Parse(teamLastActivityDate)).TotalDays;
                    }
                }

                log.LogWarning($"Unable to find team activity for {siteDisplayName}");
                return Globals.inactiveDaysDelete;
            }
            catch (Exception e)
            {
                log.LogError($"Something went wrong when trying to get team activity for {siteDisplayName} - {e.Message}");
                return Globals.inactiveDaysDelete;
            }
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
                log.LogError($"Error finding and deleting m365 group for {siteUrl} - {ex.Message}");
                success = false;
            }

            return success;
        }

        public static async Task<bool> DeleteSite(string siteUrl, ILogger log)
        {
            var success = true;

            try
            {
                var ctx = new Auth().appOnlyAuth(Globals.appOnlySiteUrl, log);
                var tenant = new Tenant(ctx);
                var removeSite = tenant.RemoveSite(siteUrl);
                ctx.Load(removeSite);
                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                // This will sometimes throw an error complaining about not being able to find a file path, but the site is successfully deleted.
                log.LogError($"Error deleting {siteUrl} - {ex}");
                success = false;
            }

            return success;
        }

        public class SiteReport
        {
            public SiteReport()
            {
                WarningSites = new List<ReportData>();
                DeleteSites = new List<ReportData>();
                NoOwnerSites = new List<ReportData>();
                StorageThresholdSites = new List<ReportData>();
            }

            public List<ReportData> WarningSites { get; set; }
            public List<ReportData> DeleteSites { get; set; }
            public List<ReportData> NoOwnerSites { get; set; }
            public List<ReportData> StorageThresholdSites { get; set; }
        }

        public class ReportData
        {
            public ReportData(string siteId, string siteUrl, int inactiveDays, List<User> siteOwners, ulong storageCapacity, ulong storageUsed)
            {
                SiteId = siteId;
                SiteUrl = siteUrl;
                InactiveDays = inactiveDays;
                SiteOwners = siteOwners;
                StorageCapacity = storageCapacity;
                StorageUsed = storageUsed;
            }

            public string SiteId { get; set; }
            public string SiteUrl { get; set; }
            public int InactiveDays { get; set; }
            public List<User> SiteOwners { get; set; }
            public ulong StorageCapacity { get; set; }
            public ulong StorageUsed { get; set; }
        }
    }
}
