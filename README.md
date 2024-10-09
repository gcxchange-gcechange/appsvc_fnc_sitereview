##  Site Review Function App
This Function App is designed to automate the monitoring and reporting of SharePoint Framework (SPFx) Teams sites across a Microsoft 365 tenant. It scans all Teams sites and generates a comprehensive report based on several criteria, such as site ownership, storage usage, activity levels, classification settings, hub association, and privacy configurations. The app aims to ensure that sites adhere to organizational policies and standards, with built-in functionality for sending alerts and automated actions when needed.
## Features
 - **Site Owners Check**: Identifies Teams sites that have fewer than a specified number of site owners.
  -   **Storage Utilization Monitoring**: Flags sites that exceed a defined percentage of their allocated storage.
-   **Activity Monitoring**:
    -   Warns when a site has been inactive for a specified number of days.
    -   Flags sites for deletion if inactivity exceeds a different, specified threshold.
-   **Classification Verification**: Checks if a site has a classification setting, and flags sites without one.
-   **Privacy Setting Validation**: Ensures that sites have the correct privacy settings based on organizational policies.
-   **Hub Association Validation**: Checks that the sites are children of a specified hub site, and flags those who are not.
-   **Automated Reporting and Notification**:
    -   Generates a report summarizing all flagged issues and emails it to a specified list of recipients.
    -   Supports "report only" mode, where no actions are taken other than sending the report.
    -   In non-report-only mode, the app sends email notifications to all site owners when a site is flagged for a warning.
    -   Automatically deletes flagged sites and associated site groups when they meet the deletion criteria.
-  **Excluded Sites**: Specific sites can be ignored from the report by providing a list of Site Ids in the app settings.
##  API Permissions
Your app registration will need the following API permissions depending on if the app is running in report only mode or not.
### Microsoft Graph
- Reports.Read.All - Read all usage reports
- Sites.Read.All - Read items in all site collections
- Group.Read.All - Read all groups
- Group.ReadWrite.All - Read and write all groups **(for report only mode OFF)**
- User.Read.All - Read all users' full profiles
- Mail.Send - Send mail as any user
### App Only
- http://sharepoint/content/sitecollection - Full Control **(for report only mode OFF)**
- http://sharepoint/content/tenant - Full Control **(for report only mode OFF)**
## How To Setup
You will need to disable the option that conceals user, group, and site names in all reports. This can be done [in the admin panel](https://admin.microsoft.com/AdminPortal/Home#/Settings/Services/:/Settings/L1/Reports)\
The function app expects the following values:
- **tenantId** - Your azure subscription
- **hubId** - The site ID for the hub site. All team sites in the tenant are expected to be a part of this hub (except the excludeSiteIds sites)
- **clientId** - The app registration client ID
- **appOnlyId** - The app only ID created in SharePoint
- **appOnlySiteUrl** - The SharePoint site url where you set up the app-principal. (eg https://{your-tenant}-admin.sharepoint.com/)
- **keyVaultUrl** - The URL to the key vault containing the client and app only secrets.
- **secretNameClient** - The name of the client secret in your key vault
- **secretNameAppOnly** - The name of the app only secret in your key vault
- **excludeSiteIds** - A string of site IDs seperated by commas. These sites will be ignored.
- **emailSenderId** - The object ID of the user that will send emails. Make sure this user has a license to send email
- **adminEmails** - The admin email addresses where the reports will be sent. Each email should be seperated by a comma.
- **inactiveDaysWarn** - The number of days of inactivity for a site to receive a warning (min 0).
- **inactiveDaysDelete** - The number of days of inactivity for a site to be deleted (min 0).
- **minSiteOwners** - The minimum number of owners for a site (min 0).
- **storageThreshold** - A number 0-100 that represents the storage used % for when a site should be flagged for review.
- **expectedPrivacySetting** - The privacy setting on sites. Anything except what's entered here will flag the site for review. (eg `Private`)
- **reportOnlyMode** - Puts the function app in report mode if set to anything but `0`.