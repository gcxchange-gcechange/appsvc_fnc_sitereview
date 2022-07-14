##  Site Review Function App
This function app looks for sites and teams that are inactive for more than 60 days or more than 120 days. Owners of the site or team will be informed after 60 days that their site or team has been flagged for deletion. After 120 days those teams or sites will be deleted and the owners will be informed. This function app runs on a monthly timer.
##  API Permissions
Your app registration will need the following API permissions
### Microsoft Graph
Reports.Read.All - Read all usage reports
Mail.Send - Send mail as any user
## How To Setup
You will need to disable the option that conceals user, group, and site names in all reports. This can be done [in the admin panel](https://admin.microsoft.com/AdminPortal/Home#/Settings/Services/:/Settings/L1/Reports)
