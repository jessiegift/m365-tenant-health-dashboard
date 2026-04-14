# M365 Tenant Health Dashboard
A PowerShell-based dashboard that polls the Microsoft Graph API 
for real-time Microsoft 365 tenant health data and publishes a live HTML dashboard to GitHub Pages,
 auto-refreshed every 6 hours via GitHub Actions.
 [View the live dashboard](https://jessiegift.github.io/m365-tenant-health-dashboard/)

  ## Architeture 
  <img width="1272" height="245" alt="Screenshot 2026-04-14 170029" src="https://github.com/user-attachments/assets/dc0f98f1-76cb-4e94-b87c-5b6d32cdf810" />


 ## What It Monitors
-Metric Graph API Endpoint 
- Service Health -admin/serviceAnnouncement/healthOverviews, Real-time status of Exchange, Teams, SharePoint etc.
- License Usage -subscribedSkus, Shows consumed vs. total licences, identifies waste 
- Inactive Users (30+ days) -users?$select=signInActivity,Security risk — stale accounts are attack vectors

## Graph API Permissions Required
This project uses an **Azure AD App Registration** with **application permissions** (no user sign-in needed):

 ## Permission  Type  Why 

- ServiceHealth.Read.All, Application, Read service health status 
- User.Read.All, Application, Read user profiles and sign-in activity
- Directory.Read.All, Application Read licence/subscription data 

## How I Set This Up
- Went to Microsoft Entra ID
- Registered a new app
- <img width="1636" height="708" alt="Screenshot 2026-04-14 122740" src="https://github.com/user-attachments/assets/f09aa8db-acbb-4ac6-98e5-8c632db2c579" />

- Granted the required Graph API permissions
- <img width="824" height="416" alt="image" src="https://github.com/user-attachments/assets/33f9c7c6-97ec-4df3-83e9-f58efd647431" />

- Created a client secret and copied the Value (not the Secret ID)
- <img width="1629" height="751" alt="image" src="https://github.com/user-attachments/assets/5d2904ca-177a-4987-aa4a-3c8ea37a1e4e" />

- Created a new GitHub repository
- Manually created the folder structure (scripts/, docs/)
- Created each file using PowerShell locally
- <img width="1181" height="913" alt="Screenshot 2026-04-14 133135" src="https://github.com/user-attachments/assets/12a9d0c6-7e05-4cd9-9d87-7041998241ac" />

- Added the script and HTML template content directly in GitHub
- Added GitHub Secrets TENANT_ID,CLIENT_ID,CLIENT_SECRET
- <img width="1900" height="939" alt="Screenshot 2026-04-14 134436" src="https://github.com/user-attachments/assets/04195cbc-a8a2-4daa-b6be-83dfd75bdf58" />
<img width="1440" height="792" alt="Screenshot 2026-04-14 134656" src="https://github.com/user-attachments/assets/adc502ee-19e3-4ce0-a6e0-4fede1418a76" />

- Enabled GitHub Pages
- <img width="1519" height="868" alt="Screenshot 2026-04-14 142203" src="https://github.com/user-attachments/assets/615cce8e-6cd7-488e-afab-dcdb71687802" />

- Source: Deploy from a branch
- Branch: main
- Folder: /docs
- Ran the GitHub Actions workflow manually
- Verified the script ran successfully
- <img width="1910" height="823" alt="Screenshot 2026-04-14 141937" src="https://github.com/user-attachments/assets/1be292f7-8548-42e3-8b77-987837ef611d" />

- Confirmed the dashboard was generated
- Fixed permissions so GitHub Actions can push updates
- Dashboard is now live and auto‑updates

