using Microsoft.Graph;
using System.Collections.Generic;

namespace SiteReview
{
    public class ReportData
    {
        public ReportData(string siteId, string siteUrl, string siteDisplayName, int inactiveDays, List<User> siteOwners, ulong storageAllocated, ulong storageUsed, string privacySetting, IEnumerable<AssignedLabel> assignedLabels, bool inHub)
        {
            SiteId = siteId;
            SiteUrl = siteUrl;
            SiteDisplayName = siteDisplayName;
            InactiveDays = inactiveDays;
            SiteOwners = siteOwners;
            StorageAllocated = storageAllocated;
            StorageUsed = storageUsed;
            PrivacySetting = privacySetting;
            AssignedLabels = assignedLabels;
            InHub = inHub;
        }

        public string SiteId { get; set; }
        public string SiteUrl { get; set; }
        public string SiteDisplayName { get; set; }
        public int InactiveDays { get; set; }
        public List<User> SiteOwners { get; set; }
        public ulong StorageAllocated { get; set; }
        public ulong StorageUsed { get; set; }
        public string PrivacySetting { get; set; }
        public IEnumerable<AssignedLabel> AssignedLabels { get; set; }
        public bool InHub { get; set; }
    }
}
