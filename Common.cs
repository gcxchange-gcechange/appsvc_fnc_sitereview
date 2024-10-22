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

                log.LogInformation($"Site usage report contains data on {siteCSV.Count - 1} sites in total.");

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
                    teamsActivityCSV = Helpers.GenerateCSV(await response.Content.ReadAsStringAsync());
                    log.LogInformation("Got teams usage report.");
                    log.LogInformation($"Teams usage report contains data on {teamsActivityCSV.Count - 1} teams in total.");
                }
                else
                {
                    log.LogError($"Error retrieving teams usage report: {response.StatusCode}");
                }

                 var allSites = await graphAPIAuth.Sites
                .Request()
                .Header("ConsistencyLevel", "eventual")
                .GetAsync();

                // Get sites in our hub
                var sitesQueryOptions = new List<QueryOption>()
                {
                    new QueryOption("search", "DepartmentId:{" + Globals.hubId + "}"),
                };

                var hubSites = await graphAPIAuth.Sites
                .Request(sitesQueryOptions)
                .Header("ConsistencyLevel", "eventual")
                .GetAsync();

                log.LogInformation($"Beginning to build your report...");

                var excludeSiteIds = Globals.GetExcludedSiteIds();

                var sitePage = 1;
                var totalSites = 0;
                do
                {
                    log.LogInformation($"{Environment.NewLine}Checking site page {sitePage} containing {allSites.Count} sites...{Environment.NewLine}");

                    foreach (var site in allSites)
                    {
                        if (excludeSiteIds.Contains(site.Id))
                        {
                            log.LogInformation($"Skipped excluded site: {site.DisplayName}");
                            continue;
                        }

                        var group = await GetGroupFromSite(site, graphAPIAuth, log);

                        if (group != null)
                        {
                            log.LogInformation($"Checking {site.DisplayName} ...");

                            // Build the report
                            for (var i = 1; i < siteCSV.Count; i++)
                            {
                                var siteId = siteCSV[i][siteSiteIdIndex];
                                var lastActivityDate = siteCSV[i][siteLastActivityIndex];
                                var siteURL = siteCSV[i][siteSiteURLIndex];
                                var storageUsed = siteCSV[i][siteStorageUsedIndex];
                                var storageAllocated = siteCSV[i][siteStorageAllocatedIndex];
                                var foundSite = site.Id.Split(",")[1] == siteId;

                                if (foundSite || i == siteCSV.Count - 1)
                                {
                                    double siteDaysInactive;

                                    if (!foundSite)
                                    {
                                        log.LogInformation($"Couldn't find {site.DisplayName} in the site usage report. Set inactive site days to {Globals.inactiveDaysDelete}.");
                                        siteDaysInactive = Globals.inactiveDaysDelete;
                                    }

                                    siteDaysInactive = lastActivityDate != String.Empty ? (DateTime.Now - DateTime.Parse(lastActivityDate)).TotalDays : Globals.inactiveDaysDelete;
                                    var teamDaysInactive = GetTeamsActivity(teamsActivityCSV, site.DisplayName, log);
                                    var siteOwners = await GetSiteOwners(site, graphAPIAuth, log);
                                    var privacySetting = group.Visibility ?? null;

                                    var reportData = new ReportData(
                                        siteId,
                                        site.WebUrl,
                                        site.DisplayName,
                                        (int)Math.Min(siteDaysInactive, teamDaysInactive),
                                        siteOwners,
                                        ulong.Parse(storageAllocated),
                                        ulong.Parse(storageUsed),
                                        privacySetting,
                                        group.AssignedLabels,
                                        SiteExists(hubSites, site)
                                    );

                                    siteReport.AddReportData(reportData);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            log.LogWarning($"Couldn't find a group for {site.DisplayName} so it will not be included in the report.");
                        }
                    }

                    sitePage++;
                    totalSites += allSites.Count;
                }
                while (allSites.NextPageRequest != null && (allSites = await allSites.NextPageRequest.GetAsync()).Count > 0);

                log.LogInformation($"{Environment.NewLine}{totalSites} sites were scanned in total.{Environment.NewLine}");

                return siteReport;
            }
            catch (Exception ex)
            {
                log.LogError($"Error building report - {ex.Message} - {ex.StackTrace}");
                return siteReport;
            }
        }

        public static bool SiteExists(IGraphServiceSitesCollectionPage sitePage, Site targetSite)
        {
            var currentPage = sitePage;

            while (currentPage != null)
            {
                if (currentPage.Contains(targetSite))
                    return true;

                
                if (currentPage.NextPageRequest != null)
                    currentPage = currentPage.NextPageRequest.GetAsync().Result;
                else
                    break;
            }

            return false;
        }

        public static async Task<Group> GetGroupFromSite(Site site, GraphServiceClient graphAPIAuth, ILogger log)
        {
            try
            {
                if (site != null)
                {
                    if (site.Name == null || site.Name == String.Empty)
                    {
                        log.LogWarning($"{site.DisplayName} will be skipped because we can't tell if it's a team or comms site.");
                        return null;
                    }

                    var escapedSiteName = site.Name.Replace(",", "%2C").Replace("&", "%26").Replace("(", "%28").Replace(")", "%29").Replace("é", "%C3%A9").Replace("É", "%C3%89").Replace(" ", "%20").Replace("'", "''");

                    var groups = await graphAPIAuth.Groups
                    .Request(new List<QueryOption>(){
                        new QueryOption("$search", "\"mailNickname:" + escapedSiteName + "\"")
                    })
                    .Header("ConsistencyLevel", "eventual")
                    .Select("id,displayName,mail,mailNickname,groupTypes,visibility,classification,assignedLicenses,assignedLabels")
                    .GetAsync();

                    if (groups != null && groups.Count > 0)
                    {
                        if (groups[0] != null)
                            return groups[0];
                    }

                    groups = await graphAPIAuth.Groups
                    .Request()
                    .Filter($"displayName eq '{escapedSiteName}'")
                    .Header("ConsistencyLevel", "eventual")
                    .Select("id,displayName,mail,mailNickname,groupTypes,visibility,classification,assignedLicenses,assignedLabels")
                    .GetAsync();

                    if (groups != null && groups.Count > 0)
                        return groups[0];
                }
            }
            catch (Exception e)
            {
                log.LogError($"Something went wrong when attempting to find the group for a site with name: {site.Name} - {e.Message}");
            }

            return null;
        }

        private static async Task<List<User>> GetSiteOwners(Site site, GraphServiceClient graphAPIAuth, ILogger log)
        {
            var siteOwners = new List<User>();
            var group = await GetGroupFromSite(site, graphAPIAuth, log);

            if (group != null)
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

            return siteOwners;
        }

        private static double GetTeamsActivity(List<List<string>> teamsActivityCSV, string siteDisplayName, ILogger log)
        {
            try
            {
                var teamNameIndex = teamsActivityCSV.FirstOrDefault().FindIndex(l => l.Equals("Team Name"));
                var lastActivityIndex = teamsActivityCSV.FirstOrDefault().FindIndex(l => l.Equals("Last Activity Date"));

                for (var i = 1; i < teamsActivityCSV.Count; i++)
                {
                    if (teamsActivityCSV[i][teamNameIndex] == siteDisplayName)
                    {
                        var teamLastActivityDate = teamsActivityCSV[i][lastActivityIndex];
                        if (teamLastActivityDate != String.Empty)
                            return (DateTime.Now - DateTime.Parse(teamLastActivityDate)).TotalDays;
                        else
                            break;
                    }
                }

                log.LogWarning($"Unable to find team activity for {siteDisplayName}. Set inactive team days to {Globals.inactiveDaysDelete}.");
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
    }
}
