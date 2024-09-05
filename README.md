##  Site Review Function App
This function app looks for sites that are inactive for more than 60 days or more than 120 days. Owners of the site or team will be informed after 60 days that their site has been flagged for deletion. After 120 days those sites will be deleted and the owners will be informed. This function app runs on a monthly timer. There are two functions, one informs owners, deletes the site teams, and stores the site ids in blob storage. The second function runs 24 hours later and deletes the sites using the ids in blob storage. The delay between the two is so there is enough time that the team is no long associated with the site and we can successfully remove the site.
##  API Permissions
Your app registration will need the following API permissions
### Microsoft Graph
- Reports.Read.All - Read all usage reports
- Sites.Read.All - Read items in all site collections
- Group.ReadWrite.All - Read and write all groups
- User.Read.All - Read all users' full profiles
- Mail.Send - Send mail as any user
### App Only
- http://sharepoint/content/sitecollection - Full Control
- http://sharepoint/content/tenant - Full Control
## How To Setup
You will need to disable the option that conceals user, group, and site names in all reports. This can be done [in the admin panel](https://admin.microsoft.com/AdminPortal/Home#/Settings/Services/:/Settings/L1/Reports)
The function app expects the following values:
- **tenantId** - Your azure subscription
- **hubId** - The site ID for the hub site. All subsites will be scanned.
- **clientId** - The app registration client ID
- **appOnlyId** - The app only ID created in SharePoint
- **appOnlySiteUrl** - The SharePoint site url where you set up the app-principal. (eg https://{your-tenant}-admin.sharepoint.com/)
- **keyVaultUrl** - The URL to the key vault containing the client and app only secrets.
- **secretNameClient** - The name of the client secret in your key vault
- **secretNameAppOnly** - The name of the app only secret in your key vault
- **excludeSiteIds** - A string of site IDs seperated by commas. These sites will be ignored.
- **emailSenderId** - The object ID of the user that will send emails. Make sure this user has a license to send email
- **adminEmail** - The admin email address where the reports will be sent
- **inactiveDaysWarn** - The number of days of inactivity for a site to receive a warning (min 0).
- **inactiveDaysDelete** - The number of days of inactivity for a site to be deleted (min 0).
- **minSiteOwners** - The minimum number of owners for a site (min 0).
- **storageThreshold** - A number 0-100 that represents when a site has reached the storage threshhold.
