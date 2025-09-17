# Office 365 License Reporting Scripts

This repository contains PowerShell scripts that generate reports from Microsoft 365:

- **O365UserLicenseReport.ps1** → Exports all users and their assigned licenses/service plans.  
- **LicenseExpiryDateReport.ps1** → Exports all subscription SKUs, license counts, and expiry dates.  

Reports are output as CSV files, which can be opened in Excel.

---

## 📋 Requirements

- **Operating System:** Windows 10 or later  
- **PowerShell Version:** Windows PowerShell 5.1 or PowerShell 7.x  
- **Microsoft 365 Account:** Global Administrator or Reports Reader role  
- **Microsoft Graph PowerShell SDK:** `Microsoft.Graph` module  

---

## 🚀 Setup

### 1. Clone or Download
Download this repository (ZIP) or clone with Git:

```bash
git clone https://github.com/<your-org>/<your-repo>.git
cd <your-repo>
# LicenceOptimization
Licence Optimization Scripts
2. Open PowerShell

Click Start → type PowerShell → right-click → Run as Administrator.

3. Allow Scripts to Run

Run this once to allow local scripts:

Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

4. Install Microsoft Graph Module

Run:

Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force


If prompted:

Install NuGet provider → press Y

Trust PSGallery → press A

▶️ Running the Scripts
A. User License Report

Exports all users and assigned licenses/service plans:

.\O365UserLicenseReport.ps1

B. License Expiry Report

Exports all subscribed SKUs with license counts and expiry dates:

.\LicenseExpiryDateReport.ps1


When prompted, sign in with your Microsoft 365 admin account.

📂 Output

Each script generates a CSV file in the same directory:

DetailedO365UserLicenseReport_YYYY-MMM-DD.csv

LicenseExpiryReport_YYYY-MMM-DD.csv

Open in Excel for filtering, sorting, and analysis.

🛠️ Troubleshooting

Script blocked:
Right-click the .ps1 file → Properties → check Unblock → OK.

No data found:
Ensure you have Global Admin or Reports Reader role in Microsoft 365.

Module not found errors:
Re-run:

Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force


Restore Execution Policy:

Set-ExecutionPolicy Restricted -Scope CurrentUser -Force

✅ Example Workflow
# Navigate to script folder
cd "C:\O365Reports"

# Run user license report
.\O365UserLicenseReport.ps1

# Run subscription expiry report
.\LicenseExpiryDateReport.ps1
