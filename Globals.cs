using System;
using System.Collections.Generic;
using System.Linq;

namespace SiteReview
{
    public static class Globals
    {
        // TODO: Add the expected classifications as an app setting once we know which classifications we want to enforce.
        //       For now the app will flag anything without a classification.

        public const string privateSetting = "Private";
        public static readonly string tenantId = GetEnvironmentString("tenantId");
        public static readonly string clientId = GetEnvironmentString("clientId");
        public static readonly string hubId = GetEnvironmentString("hubId");
        public static readonly string[] adminEmails = GetEnvironmentString("adminEmails").Split(',').Select(Email => Email.Trim()).ToArray();
        public static readonly string emailUserName = GetEnvironmentString("emailUserName");

        public static readonly string keyVaultUrl = GetEnvironmentString("keyVaultUrl");
        public static readonly string secretNameClient = GetEnvironmentString("secretNameClient");

        public static readonly string appOnlySiteUrl = GetEnvironmentString("appOnlySiteUrl");
        public static readonly string appOnlyId = GetEnvironmentString("appOnlyId");
        public static readonly string secretNameAppOnly = GetEnvironmentString("secretNameAppOnly");

        public static readonly int inactiveDaysWarn = GetEnvironmentInt("inactiveDaysWarn", 0);
        public static readonly int inactiveDaysDelete = GetEnvironmentInt("inactiveDaysDelete", 0);
        public static readonly int minSiteOwners = GetEnvironmentInt("minSiteOwners", 0);
        public static readonly double storageThreshold = GetEnvironmentDouble("storageThreshold", 0, 100);
        public static readonly string expectedPrivacySetting = GetEnvironmentString("expectedPrivacySetting");

        public static readonly bool reportOnlyMode = GetEnvironmentString("reportOnlyMode") != "0";

        public static List<string> GetExcludedSiteIds()
        {
            var excludedSiteIds = new List<string>(GetEnvironmentString("excludeSiteIds").Replace(" ", "").Split(","));
            excludedSiteIds.Add(hubId);

            return excludedSiteIds;
        }

        private static string GetEnvironmentString(string name)
        {
            return Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }

        private static int GetEnvironmentInt(string name, int min = int.MinValue, int max = int.MaxValue)
        {
            var numericString = GetEnvironmentString(name);
            int retVal;

            if (int.TryParse(numericString, out retVal))
            {
                retVal = Math.Clamp(retVal, min, max);
            }

            return retVal;
        }

        private static double GetEnvironmentDouble(string name, double min = double.MinValue, double max = double.MaxValue)
        {
            var numericString = GetEnvironmentString(name);
            double retVal;

            if (double.TryParse(numericString, out retVal))
            {
                retVal = Math.Clamp(retVal, min, max);
            }

            return retVal;
        }
    }
}

