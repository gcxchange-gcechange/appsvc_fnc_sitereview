##  Site Review Function App
This function app looks for sites and teams that are inactive for more than 60 days or more than 120 days. Owners of the site or team will be informed after 60 days that their site or team has been flagged for deletion. After 120 days those teams or sites will be deleted and the owners will be informed. This function app runs on a monthly timer.
##  API Permissions
Your app registration will need the following API permissions
### Microsoft Graph
- Reports.Read.All - Read all usage reports
- Sites.Read.All - Read items in all site collections
- Group.ReadWrite.All - Read and write all groups
- Mail.Send - Send mail as any user
### App Only
http://sharepoint/content/sitecollection - Full Control
http://sharepoint/content/tenant - Full Control
## How To Setup
You will need to disable the option that conceals user, group, and site names in all reports. This can be done [in the admin panel](https://admin.microsoft.com/AdminPortal/Home#/Settings/Services/:/Settings/L1/Reports)
The function app expects the following values:
- **tenantId** - Your azure subscription
- **hubId** - The site ID for the hub site. All subsites will be scanned.
- **clientId** - The app registration client ID
- **appOnlyId** - The app only ID created in SharePoint
- **keyVaultUrl** - The URL to the key vault containing the client and app only secrets.
- **secretNameClient** - The name of the client secret in your key vault
- **secretNameAppOnly** - The name of the app only secret in your key vault
- **excludeSiteIds** - A string of site IDs seperated by commas. These sites will be ignored.
- **emailSenderId** - The object ID of the user that will send emails. Make sure this user has a license to send email.