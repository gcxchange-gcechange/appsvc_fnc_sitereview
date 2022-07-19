using System.Collections.Generic;

namespace SiteReview
{
    public static class Globals
    {
        public static readonly string tenantId = GetEnvironmentVariable("tenantId");
        public static readonly string clientId = GetEnvironmentVariable("clientId");
        public static readonly string hubId = GetEnvironmentVariable("hubId");
        public static readonly string emailSenderId = GetEnvironmentVariable("emailSenderId");

        public static readonly string keyVaultUrl = GetEnvironmentVariable("keyVaultUrl");
        public static readonly string secretNameClient = GetEnvironmentVariable("secretNameClient");

        public static readonly string appOnlyId = GetEnvironmentVariable("appOnlyId");
        public static readonly string secretNameAppOnly = GetEnvironmentVariable("secretNameAppOnly");

        public static List<string> excludeTeamIds = new List<string>(GetEnvironmentVariable("excludeTeamIds").Replace(" ", "").Split(","));

        public static List<string> GetExcludedSiteIds()
        {
            var excludedSiteIds = new List<string>(GetEnvironmentVariable("excludeSiteIds").Replace(" ", "").Split(","));
            excludedSiteIds.Add(hubId);

            return excludedSiteIds;
        }

        private static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, System.EnvironmentVariableTarget.Process);
        }
    }
}

