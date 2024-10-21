using Microsoft.Graph;
using System.Collections.Generic;

namespace SiteReview
{
    public class ReportData
    {
        public ReportData(string siteId, string siteUrl, string siteDisplayName, int inactiveDays, List<User> siteOwners, ulong storageCapacity, ulong storageUsed, string privacySetting, string classification, bool inHub)
        {
            SiteId = siteId;
            SiteUrl = siteUrl;
            SiteDisplayName = siteDisplayName;
            InactiveDays = inactiveDays;
            SiteOwners = siteOwners;
            StorageCapacity = storageCapacity;
            StorageUsed = storageUsed;
            PrivacySetting = privacySetting;
            Classification = classification;
            InHub = inHub;
        }

        public string SiteId { get; set; }
        public string SiteUrl { get; set; }
        public string SiteDisplayName { get; set; }
        public int InactiveDays { get; set; }
        public List<User> SiteOwners { get; set; }
        public ulong StorageCapacity { get; set; }
        public ulong StorageUsed { get; set; }
        public string PrivacySetting { get; set; }
        public string Classification { get; set; }
        public bool InHub { get; set; }
    }
}
