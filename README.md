# M365 Tenant Health Dashboard
A PowerShell-based dashboard that polls the Microsoft Graph API 
for real-time Microsoft 365 tenant health data and publishes a live HTML dashboard to GitHub Pages,
 auto-refreshed every 6 hours via GitHub Actions.
 [View the live dashboard](https://jessiegift.github.io/m365-tenant-health-dashboard/)

 ## What It Monitors
-Metric Graph API Endpoint 
- Service Health -admin/serviceAnnouncement/healthOverviews, Real-time status of Exchange, Teams, SharePoint etc.
- License Usage -subscribedSkus, Shows consumed vs. total licences, identifies waste 
- Inactive Users (30+ days) -users?$select=signInActivity,Security risk — stale accounts are attack vectors

## Graph API Permissions Required
This project uses an **Azure AD App Registration** with **application permissions** (no user sign-in needed):

| Permission | Type | Why |

ServiceHealth.Read.All, Application, Read service health status 
User.Read.All, Application, Read user profiles and sign-in activity
Directory.Read.All, Application Read licence/subscription data 

 Client secret is stored in **GitHub Secrets**.
How I Set This Up
Went to Microsoft Entra ID
Registered a new app
Granted the required Graph API permissions
Created a client secret and copied the Value (not the Secret ID)
Created a new GitHub repository
Manually created the folder structure (scripts/, docs/)
Created each file using PowerShell locally
Added the script and HTML template content directly in GitHub
Added GitHub Secrets
TENANT_ID,CLIENT_ID,CLIENT_SECRET
Enabled GitHub Pages
Source: Deploy from a branch
Branch: main
Folder: /docs
Ran the GitHub Actions workflow manually
Verified the script ran successfully
Confirmed the dashboard was generated
Fixed permissions so GitHub Actions can push updates
Dashboard is now live and auto‑updates

