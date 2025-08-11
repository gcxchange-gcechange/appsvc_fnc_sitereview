using AngleSharp.Dom;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.IO;
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
            var siteReport = new SiteReport(log);

            try
            {
                // Get a report of site usage in the last 180 days
                var requestInfo = new RequestInformation
                {
                    HttpMethod = Method.GET,
                    UrlTemplate = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D180')",
                    PathParameters = new Dictionary<string, object>(),
                };
                requestInfo.Headers.Add("ConsistencyLevel", "eventual");

                var siteReportResponse = await graphAPIAuth.RequestAdapter.SendPrimitiveAsync<Stream>(requestInfo, cancellationToken: default);

                log.LogInformation("Got site usage report.");

                using var reader = new StreamReader(siteReportResponse);

                // Download the site CSV data
                var siteCSV = Helpers.GenerateCSV(await reader.ReadToEndAsync());

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

                var allSites = await graphAPIAuth.Sites.GetAsync(requestConfig =>
                {
                    requestConfig.Headers.Add("ConsistencyLevel", "eventual");
                });

                // Get sites in our hub
                var hubSitesPage = await graphAPIAuth.Sites[Globals.hubId].Sites.GetAsync(requestConfig =>
                {
                    requestConfig.Headers.Add("ConsistencyLevel", "eventual");
                });

                var hubSites = new List<Site>();

                while (hubSitesPage != null)
                {
                    hubSites.AddRange(hubSitesPage.Value);

                    if (string.IsNullOrEmpty(hubSitesPage.OdataNextLink))
                    {
                        break;
                    }

                    var nextRequestBuilder = new Microsoft.Graph.Sites.SitesRequestBuilder(hubSitesPage.OdataNextLink, graphAPIAuth.RequestAdapter);
                    hubSitesPage = await nextRequestBuilder.GetAsync();
                }

                log.LogInformation($"Beginning to build your report...");

                var excludeSiteIds = Globals.GetExcludedSiteIds();

                var sitePage = 1;
                var totalSites = 0;
                var teamsSites = 0;

                while (allSites != null)
                {
                    log.LogInformation($"{Environment.NewLine}Checking site page {sitePage} containing {allSites.Value.Count} sites...{Environment.NewLine}");

                    foreach (var site in allSites.Value)
                    {
                        if (excludeSiteIds.Contains(site.Id.Split(",")[1]))
                        {
                            log.LogInformation($"Skipped excluded site: {site.DisplayName}");
                            continue;
                        }

                        var group = await GetGroupFromSite(site, graphAPIAuth, log);

                        if (group != null)
                        {
                            log.LogInformation($"{Environment.NewLine}Checking {site.DisplayName} ...");
                            teamsSites++;

                            // Build the report
                            for (var i = 1; i < siteCSV.Count; i++)
                            {
                                var siteId = siteCSV[i][siteSiteIdIndex];
                                var lastActivityDate = siteCSV[i][siteLastActivityIndex];
                                var storageUsed = siteCSV[i][siteStorageUsedIndex];
                                var storageAllocated = siteCSV[i][siteStorageAllocatedIndex];
                                var foundSite = site.Id.Split(",")[1] == siteId;

                                if (foundSite || i == siteCSV.Count - 1)
                                {
                                    double siteDaysInactive;

                                    if (!foundSite)
                                    {
                                        log.LogWarning($"Couldn't find {site.DisplayName} in the site usage report. Set inactive site days to {Globals.inactiveDaysDelete}.");
                                        siteDaysInactive = Globals.inactiveDaysDelete;
                                    }
                                    else
                                    {
                                        log.LogInformation(Environment.NewLine +
                                            $"Site activity report for {site.DisplayName}{Environment.NewLine}" +
                                            $"siteId: {siteCSV[i][siteSiteIdIndex]}{Environment.NewLine}" +
                                            $"lastActivityDate: {siteCSV[i][siteLastActivityIndex]}{Environment.NewLine}" +
                                            $"storageAllocated: {siteCSV[i][siteStorageAllocatedIndex]} Bytes{Environment.NewLine}" +
                                            $"storageUsed: {siteCSV[i][siteStorageUsedIndex]} Bytes"
                                        );

                                        siteDaysInactive = lastActivityDate != String.Empty ? (DateTime.Now - DateTime.Parse(lastActivityDate)).TotalDays : Globals.inactiveDaysDelete;

                                        if (lastActivityDate == string.Empty)
                                            log.LogWarning($"Unable to find site activity for {site.DisplayName}. Set inactive site days to 120.");
                                    }

                                    var teamDaysInactive = GetTeamsActivity(teamsActivityCSV, site.DisplayName, log);
                                    var siteOwners = await GetSiteOwners(site, graphAPIAuth, log, group);
                                    var privacySetting = group.Visibility ?? null;

                                    var reportData = new ReportData(
                                        site.Id.Split(",")[1],
                                        site.WebUrl,
                                        site.DisplayName,
                                        (int)Math.Min(siteDaysInactive, teamDaysInactive),
                                        siteOwners,
                                        ulong.Parse(foundSite ? storageAllocated : "0"),
                                        ulong.Parse(foundSite ? storageUsed : "0"),
                                        privacySetting,
                                        group.AssignedLabels,
                                        hubSites.Any(s => s.Id == site.Id)
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
                    totalSites += allSites.Value.Count;

                    if (string.IsNullOrEmpty(allSites.OdataNextLink))
                    {
                        allSites = null;
                    }
                    else
                    {
                        var nextRequestBuilder = new Microsoft.Graph.Sites.SitesRequestBuilder(allSites.OdataNextLink, graphAPIAuth.RequestAdapter);
                        allSites = await nextRequestBuilder.GetAsync();
                    }
                }

                log.LogInformation($"{Environment.NewLine}{totalSites} sites were scanned in total.");
                log.LogInformation($"{teamsSites} of those were found to be teams sites.");
                log.LogInformation($"{siteReport.GetUniqueListSites().Count} of those were in violation of one or more of our policies.{Environment.NewLine}");

                return siteReport;
            }
            catch (Exception ex)
            {
                log.LogError($"Error building report - {ex.Message} - {ex.StackTrace}");
                return siteReport;
            }
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
                    if (escapedSiteName.Length > 100)
                        escapedSiteName = escapedSiteName.Substring(0, 99);

                    var groupsResponse = await graphAPIAuth.Groups.GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = $"startswith(displayName,'{escapedSiteName}')";
                        requestConfig.Headers.Add("ConsistencyLevel", "eventual");
                        requestConfig.QueryParameters.Select = new[]
                        {
                            "id",
                            "displayName",
                            "mail",
                            "mailNickname",
                            "groupTypes",
                            "visibility",
                            "classification",
                            "assignedLicenses",
                            "assignedLabels"
                        };
                    });

                    while (groupsResponse != null)
                    {
                        foreach (var group in groupsResponse.Value)
                        {
                            if (group.DisplayName == site.Name)
                                return group;
                        }

                        if (string.IsNullOrEmpty(groupsResponse.OdataNextLink))
                        {
                            break;
                        }

                        var nextRequestBuilder = new Microsoft.Graph.Groups.GroupsRequestBuilder(groupsResponse.OdataNextLink, graphAPIAuth.RequestAdapter);
                        groupsResponse = await nextRequestBuilder.GetAsync();
                    }
                }
            }
            catch (Exception e)
            {
                log.LogError($"Something went wrong when attempting to find the group for a site with name: {site.Name} - {e.Message}");
            }

            return null;
        }

        private static async Task<List<User>> GetSiteOwners(Site site, GraphServiceClient graphAPIAuth, ILogger log, Group group = null)
        {
            var siteOwners = new List<User>();

            if (group == null)
                group = await GetGroupFromSite(site, graphAPIAuth, log);

            try
            {
                if (group != null)
                {
                    var ownersResponse = await graphAPIAuth.Groups[group.Id].Owners.GetAsync();
                    var owners = ownersResponse.Value;

                    while (owners != null)
                    {
                        foreach (var owner in owners)
                        {
                            if (owner is User userOwner)
                            {
                                var user = await graphAPIAuth.Users[userOwner.Id]
                                   .GetAsync(requestConfig =>
                                   {
                                       requestConfig.QueryParameters.Select = new[] { "displayName", "mail" };
                                   });

                                if (user != null)
                                {
                                    siteOwners.Add(user);
                                }
                            }
                        }

                        if (string.IsNullOrEmpty(ownersResponse.OdataNextLink))
                        {
                            break;
                        }

                        var nextRequestBuilder = new Microsoft.Graph.Groups.Item.Owners.OwnersRequestBuilder(ownersResponse.OdataNextLink, graphAPIAuth.RequestAdapter);
                        ownersResponse = await nextRequestBuilder.GetAsync();
                        owners = ownersResponse.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error retrieving site owners for {site.DisplayName} - {ex.Message} - {ex.StackTrace}");
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
                        log.LogInformation(Environment.NewLine +
                            $"Teams activity report for {siteDisplayName}{Environment.NewLine}" +
                            $"lastActivityDate: {teamsActivityCSV[i][lastActivityIndex]}"
                        );

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

                await graphAPIAuth.Groups[groupId.ToString()].DeleteAsync();
            }
            catch (Exception ex)
            {
                log.LogError($"Error finding and deleting m365 group for {siteUrl} - {ex.Message}");
                success = false;
            }

            return success;
        }

        public static bool DeleteSite(string siteUrl, ILogger log)
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
