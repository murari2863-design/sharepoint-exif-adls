# SharePoint Image Metadata Exporter → Azure Data Lake Storage Gen2

A reusable Python utility and Azure Function that scans a SharePoint document-library folder recursively through Microsoft Graph and exports metadata for **verified images only**.

Generated Excel, CSV, JSON, debug data, and checkpoint files are stored in **Azure Data Lake Storage Gen2 (ADLS Gen2)** so multiple authorized users, virtual machines, and Azure Functions can share the same output and resume state.

The exporter is metadata-only. Image binaries are **not downloaded** from SharePoint. It reads Microsoft Graph image/GPS/photo metadata and SharePoint verification fields, then persists the results to ADLS.

The source workflow marks a file processed only after metadata retrieval, verification, and export processing complete, allowing failures to be retried safely.

---

# Architecture

```text
SharePoint
    │
    │ Microsoft Graph
    ▼
Azure Function App
    │
    │ Managed Identity
    ▼
Azure Data Lake Storage Gen2
    │
    ├── sharepoint-image-metadata-exporter/
    │     ├── excel/image_metadata.xlsx
    │     ├── csv/image_metadata.csv
    │     ├── json/image_metadata.json
    │     ├── debug/debug.csv
    │     └── checkpoint/checkpoint.json
    │
    ▼
Power BI
```

The Azure Function runs on a schedule, scans SharePoint through Microsoft Graph, extracts metadata for verified images, and stores all outputs in ADLS Gen2.

The local cache exists only during execution.

The authoritative copies are stored in ADLS.

The checkpoint is uploaded after each successfully processed image.

---

# Features

- Recursive SharePoint folder scanning
- Microsoft Graph integration
- Verified-image filtering
- GPS metadata extraction
- Camera metadata extraction
- Image dimension extraction
- Excel export
- CSV export
- JSON export
- Debug CSV export
- Resume support using checkpoints
- Azure Data Lake Storage Gen2 integration
- Azure Function support
- Managed Identity authentication
- Power BI friendly outputs

---

# Requirements

## Local Development

- Python 3.11+
- Azure CLI
- Azure Data Lake Storage Gen2 account with hierarchical namespace enabled
- Microsoft 365 / SharePoint access
- GitHub account

## Azure Deployment

- Azure Function App (Python)
- System Assigned Managed Identity enabled
- Azure Data Lake Storage Gen2 account
- SharePoint site access
- Microsoft Graph permissions

---

# Repository Structure

```text
sharepoint-exif-adls
│
├── host.json
├── function_app.py
├── requirements.txt
├── README.md
├── .github/
│   └── workflows/
│
└── output/
```

Root-level files are required for Azure Function deployment.

---

# host.json

```json
{
  "version": "2.0"
}
```

---

# Local Setup

## 1. Clone

```bash
git clone <YOUR_GITHUB_REPOSITORY_URL>
cd sharepoint-exif-adls
```

---

## 2. Create Virtual Environment

### Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### macOS/Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

---

## 3. Install Dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

---

## 4. Configure Environment Variables

Create a `.env` file.

Example:

```text
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

SHAREPOINT_SITE_URL=https://corpdir.sharepoint.com/sites/DWT_PlantPlanning_566

SHAREPOINT_FOLDER_PATH=Construction Photos

SHAREPOINT_LIBRARY_CANDIDATES=Documents,Shared Documents

VERIFIED_COLUMN_DISPLAY_NAME=Verified By Uploader

ADLS_ACCOUNT_URL=https://images566.dfs.core.windows.net

ADLS_FILE_SYSTEM=images-exif

ADLS_OUTPUT_PREFIX=sharepoint-image-metadata-exporter

DEBUG=false

RESET_CHECKPOINT=false
```

Do not commit `.env`.

---

## 5. Authenticate

For local development:

```bash
az login --tenant YOUR_TENANT_ID
```

The exporter uses the existing Azure CLI identity.

---

## 6. Grant Storage Access

Grant the running identity:

```text
Storage Blob Data Contributor
```

on:

```text
images566
```

Without data-plane access, uploads to ADLS will fail.

---

## 7. Run Locally

```bash
python function_app.py
```

---

# Azure Function Deployment

The solution can be deployed directly from GitHub Actions.

Example GitHub workflow settings:

```yaml
env:
  AZURE_FUNCTIONAPP_PACKAGE_PATH: '.'
  PYTHON_VERSION: '3.11'
```

---

# Azure Function Configuration

Enable:

```text
Function App
→ Identity
→ System Assigned
→ On
```

Save the configuration.

---

# Function App Environment Variables

Go to:

```text
Function App
→ Settings
→ Environment Variables
```

Configure the following:

```text
SHAREPOINT_SITE_URL
SHAREPOINT_FOLDER_PATH
SHAREPOINT_LIBRARY_CANDIDATES
VERIFIED_COLUMN_DISPLAY_NAME

ADLS_ACCOUNT_URL
ADLS_FILE_SYSTEM
ADLS_OUTPUT_PREFIX

OUTPUT_EXCEL
OUTPUT_CSV
OUTPUT_JSON

DEBUG_CSV
CHECKPOINT_FILE

DEBUG
RESET_CHECKPOINT

TIMER_SCHEDULE
```

Example:

```text
SHAREPOINT_SITE_URL=https://corpdir.sharepoint.com/sites/DWT_PlantPlanning_566

SHAREPOINT_FOLDER_PATH=Construction Photos

SHAREPOINT_LIBRARY_CANDIDATES=Documents,Shared Documents

VERIFIED_COLUMN_DISPLAY_NAME=Verified By Uploader

ADLS_ACCOUNT_URL=https://images566.dfs.core.windows.net

ADLS_FILE_SYSTEM=images-exif

ADLS_OUTPUT_PREFIX=sharepoint-image-metadata-exporter

TIMER_SCHEDULE=0 0 */6 * * *
```

---

# Authentication

## Local Development

Uses:

```python
DefaultAzureCredential()
```

which automatically uses:

```bash
az login
```

authentication.

---

## Azure Function

Uses:

```python
DefaultAzureCredential()
```

with:

```text
Managed Identity
```

No secrets are required.

No CLIENT_ID is required.

No CLIENT_SECRET is required.

No credentials are stored in source control.

---

# Azure Data Lake Storage Permissions

Grant the Function App Managed Identity:

```text
Storage Blob Data Contributor
```

on:

```text
Storage Account: images566
```

Without this role assignment the exporter cannot:

- upload Excel files
- upload CSV files
- upload JSON files
- upload debug files
- upload checkpoints

---

# SharePoint Permissions

The Azure administrator does not need to be a SharePoint site member.

The Function App identity requires SharePoint access.

Recommended:

```text
Sites.Selected
```

Then grant access only to:

```text
https://corpdir.sharepoint.com/sites/DWT_PlantPlanning_566
```

Alternative:

```text
Sites.Read.All
Files.Read.All
```

if organizational security policy allows broader access.

---

# Output Layout

```text
images-exif
└── sharepoint-image-metadata-exporter/
    ├── excel/
    │   └── image_metadata.xlsx
    │
    ├── csv/
    │   └── image_metadata.csv
    │
    ├── json/
    │   └── image_metadata.json
    │
    ├── debug/
    │   └── debug.csv
    │
    └── checkpoint/
        └── checkpoint.json
```

---

# Why the Checkpoint is Stored in ADLS

The checkpoint contains processed SharePoint File IDs.

Benefits:

- Resume interrupted scans
- Share progress between machines
- Share progress between Azure Functions
- Avoid re-reading already processed images
- Support scheduled execution

---

# Resetting a Scan

To start over:

```text
RESET_CHECKPOINT=true
CLEAN_OUTPUT_ON_RESET=true
```

This removes:

```text
Excel
CSV
JSON
Debug
Checkpoint
```

artifacts and creates a fresh export.

Use carefully.

---

# Configuration Reference

| Variable | Required | Description |
|-----------|-----------|-------------|
| AZURE_TENANT_ID | Yes | Microsoft Entra tenant ID |
| SHAREPOINT_SITE_URL | Yes | SharePoint site URL |
| SHAREPOINT_FOLDER_PATH | Yes | Folder relative to library root |
| SHAREPOINT_LIBRARY_CANDIDATES | No | Library names to try |
| VERIFIED_COLUMN_DISPLAY_NAME | No | Verification column display name |
| ADLS_ACCOUNT_URL | Yes | ADLS DFS endpoint |
| ADLS_FILE_SYSTEM | Yes | ADLS container/file system |
| ADLS_OUTPUT_PREFIX | No | Output root folder |
| ADLS_LOCAL_CACHE | No | Temporary local cache |
| OUTPUT_EXCEL | No | Excel path |
| OUTPUT_CSV | No | CSV path |
| OUTPUT_JSON | No | JSON path |
| DEBUG_CSV | No | Debug CSV path |
| CHECKPOINT_FILE | No | Checkpoint path |
| SSL_VERIFY | No | Default true |
| DEBUG | No | Default false |
| EXCEL_SAVE_INTERVAL | No | Save interval |
| TIMER_SCHEDULE | No | Azure Function schedule |

---

# Security

- No client secrets stored in source control
- No passwords stored in source control
- Managed Identity supported
- Azure CLI authentication supported
- Image files are not downloaded
- ADLS access uses Azure RBAC
- Generated outputs can be isolated using separate prefixes

Before publishing publicly review:

- Tenant IDs
- SharePoint URLs
- Storage account names
- Storage paths
- Internal SharePoint field names
- Sample output files
- Logs

---

# Troubleshooting

## Deployment Succeeds but No Functions Appear

Verify:

```text
host.json
function_app.py
requirements.txt
```

exist in the deployment root.

---

## Trigger Synchronization Failed (502)

Typical causes:

```text
Missing dependency
Python startup failure
Function indexing failure
Syntax error
```

Check:

```text
Function App
→ Monitoring
→ Log Stream
```

and review startup logs.

---

## No Functions Found

Ensure:

```text
function_app.py
```

contains:

```python
app = func.FunctionApp()
```

and at least one trigger:

```python
@app.timer_trigger(...)
```

---

## ADLS Access Error

Example:

```text
AuthorizationPermissionMismatch
```

Grant:

```text
Storage Blob Data Contributor
```

to the Function App Managed Identity.

---

## SharePoint Access Error

Example:

```text
403 Forbidden
```

or

```text
Insufficient privileges
```

Verify:

- Graph permissions are granted
- SharePoint site access is granted
- Managed Identity is authorized

---

## Storage Account URL

Use:

```text
https://images566.dfs.core.windows.net
```

Do not use:

```text
https://images566.blob.core.windows.net
```

for `ADLS_ACCOUNT_URL`.

---

## Concurrent Execution

Do not run multiple exporters using the same:

```text
ADLS_OUTPUT_PREFIX
```

at the same time.

Use separate prefixes per worker and merge outputs later.

---

# License

Choose a license appropriate for your organization before publishing publicly.
