using System.Collections.Generic;
using System.Linq;

namespace SiteReview
{
    public class SiteReport
    {
        public SiteReport()
        {
            WarningSites = new List<ReportData>();
            DeleteSites = new List<ReportData>();
            NoOwnerSites = new List<ReportData>();
            StorageThresholdSites = new List<ReportData>();
            PrivacySettingSites = new List<ReportData>();
            ClassificationSites = new List<ReportData>();
            HubAssociationSites = new List<ReportData>();
        }

        public List<ReportData> WarningSites { get; set; }
        public List<ReportData> DeleteSites { get; set; }
        public List<ReportData> NoOwnerSites { get; set; }
        public List<ReportData> StorageThresholdSites { get; set; }
        public List<ReportData> PrivacySettingSites { get; set; }
        public List<ReportData> ClassificationSites { get; set; }
        public List<ReportData> HubAssociationSites { get; set; }

        public void AddReportData(ReportData reportData)
        {
            if (reportData.InactiveDays >= Globals.inactiveDaysDelete)
                DeleteSites.Add(reportData);
            else if (reportData.InactiveDays >= Globals.inactiveDaysWarn)
                WarningSites.Add(reportData);

            if (reportData.SiteOwners.Count < Globals.minSiteOwners)
                NoOwnerSites.Add(reportData);

            var usedPercentage = reportData.StorageUsed / reportData.StorageCapacity * 100;
            if (usedPercentage >= Globals.storageThreshold)
                StorageThresholdSites.Add(reportData);

            if (reportData.PrivacySetting != Globals.expectedPrivacySetting)
                PrivacySettingSites.Add(reportData);

            if (!reportData.AssignedLabels.Any())
                ClassificationSites.Add(reportData);

            if (!reportData.InHub)
                HubAssociationSites.Add(reportData);
        }

        public List<ReportData> GetUniqueListSites()
        {
            return WarningSites
                .Concat(DeleteSites)
                .Concat(NoOwnerSites)
                .Concat(StorageThresholdSites)
                .Concat(PrivacySettingSites)
                .Concat(ClassificationSites)
                .GroupBy(site => site.SiteId)
                .Select(site => site.First())
                .ToList();
        }
    }
}
