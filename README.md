# SharePoint Image Metadata Exporter → Azure Data Lake Storage Gen2

A reusable Python utility that scans a SharePoint document-library folder recursively through Microsoft Graph and exports metadata for **verified images only**. Generated Excel, CSV, JSON, debug data, and checkpoint files are stored in **Azure Data Lake Storage Gen2 (ADLS Gen2)** so multiple authorized users/machines can share the same output and resume state.

The exporter is metadata-only: image binaries are not downloaded from SharePoint. It reads Graph image/GPS/photo metadata and SharePoint verification fields, then persists the results to ADLS. The source workflow marks a file processed only after metadata retrieval, verification, and the export decision complete, so failures can be retried on a later run.

## Architecture

```text
SharePoint
   │
   │ Microsoft Graph
   ▼
Python exporter
   │
   │ Azure identity
   ▼
Azure Data Lake Storage Gen2
   │
   ├── sharepoint-image-metadata-exporter/
   │   ├── excel/image_metadata.xlsx
   │   ├── csv/image_metadata.csv
   │   ├── json/image_metadata.json
   │   ├── debug/debug.csv
   │   └── checkpoint/checkpoint.json
```

The files are cached locally only while the program is running; the authoritative copies are in ADLS. The checkpoint is uploaded after each successfully processed image.

## Requirements

- Python 3.10+
- Azure CLI
- Azure Data Lake Storage Gen2 account with hierarchical namespace enabled
- Microsoft 365/SharePoint access to the source library
- Azure permissions to write/read the configured ADLS file system, normally **Storage Blob Data Contributor**
- Microsoft Graph access sufficient to read the SharePoint site, library, items and fields

## Setup

### 1. Clone

```bash
git clone <YOUR_GITHUB_REPOSITORY_URL>
cd sharepoint-image-metadata-exporter
```

### 2. Virtual environment

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Configure

Copy `.env.example` to `.env` and set:

```text
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
SHAREPOINT_SITE_URL=https://contoso.sharepoint.com/sites/Engineering
SHAREPOINT_FOLDER_PATH=General/Projects/Progress Photos

ADLS_ACCOUNT_URL=https://mystorageaccount.dfs.core.windows.net
ADLS_FILE_SYSTEM=metadata
ADLS_OUTPUT_PREFIX=sharepoint-image-metadata-exporter
```

Do not commit `.env`.

### 5. Sign in

```bash
az login --tenant YOUR_TENANT_ID
```

The exporter uses the existing Azure CLI identity for both Microsoft Graph and ADLS.

### 6. Give the user ADLS access

Grant the user or group running the exporter an appropriate **data-plane** role on the storage account/container, typically `Storage Blob Data Contributor`. An ordinary Azure management role such as Reader is not enough to upload files.

### 7. Run

```bash
python src/sharepoint_image_metadata_exporter.py
```

## ADLS output layout

With the example configuration, the program creates:

```text
metadata/
└── sharepoint-image-metadata-exporter/
    ├── excel/image_metadata.xlsx
    ├── csv/image_metadata.csv
    ├── json/image_metadata.json
    ├── debug/debug.csv
    └── checkpoint/checkpoint.json
```

Change `ADLS_OUTPUT_PREFIX` if several independent SharePoint scans need separate output areas.

## Why the checkpoint is in ADLS

The checkpoint contains the SharePoint file IDs that have already completed processing. Because it is stored remotely, a second authorized machine can download the same checkpoint and continue the scan instead of starting from zero. This also makes the solution suitable for scheduled jobs or a VM/container later.

**Do not run two exporters against the same output prefix at the same time.** They could overwrite each other's checkpoint/output. If parallel processing is required, use a separate prefix per worker and merge results afterward.

## Reset

To start a completely new scan:

```text
RESET_CHECKPOINT=true
CLEAN_OUTPUT_ON_RESET=true
```

This removes the configured Excel, CSV, JSON, debug and checkpoint artifacts from ADLS and starts again. Use this carefully.

## Configuration

| Variable | Required | Description |
|---|---:|---|
| `AZURE_TENANT_ID` | Yes | Microsoft Entra tenant ID |
| `SHAREPOINT_SITE_URL` | Yes | SharePoint site URL |
| `SHAREPOINT_FOLDER_PATH` | Yes | Folder relative to library root |
| `SHAREPOINT_LIBRARY_CANDIDATES` | No | Library names to try |
| `VERIFIED_COLUMN_DISPLAY_NAME` | No | Verification column display name |
| `ADLS_ACCOUNT_URL` | Yes | ADLS Gen2 DFS endpoint |
| `ADLS_FILE_SYSTEM` | Yes | ADLS container/file system |
| `ADLS_OUTPUT_PREFIX` | No | Root folder for this export |
| `ADLS_LOCAL_CACHE` | No | Temporary local cache |
| `OUTPUT_EXCEL` | No | Excel path inside prefix |
| `OUTPUT_CSV` | No | CSV path inside prefix |
| `OUTPUT_JSON` | No | JSON path inside prefix |
| `DEBUG_CSV` | No | Debug CSV path inside prefix |
| `CHECKPOINT_FILE` | No | Checkpoint path inside prefix |
| `SSL_VERIFY` | No | Defaults to true |
| `DEBUG` | No | Defaults to false |
| `EXCEL_SAVE_INTERVAL` | No | Periodic Excel/CSV/JSON save interval |

## Security

- No client secret is stored in the repository.
- Azure CLI authentication is used for the signed-in user.
- `.env`, tokens, generated files and local cache are gitignored.
- SharePoint image binaries are not downloaded.
- ADLS access uses Azure identity and data-plane RBAC.

Before publishing publicly, review the repository for organization-specific URLs, field names, tenant IDs, sample data, and logs.

## Troubleshooting

### ADLS authentication/authorization fails

Run:

```bash
az login --tenant YOUR_TENANT_ID
az account show
```

Then verify the identity has a data-plane role such as `Storage Blob Data Contributor` on the storage account or file system.

### Storage account URL

Use the DFS endpoint:

```text
https://STORAGE_ACCOUNT_NAME.dfs.core.windows.net
```

Do not use the blob endpoint in `ADLS_ACCOUNT_URL`.

### SharePoint scan works but output upload fails

The Graph permissions and ADLS permissions are independent. Confirm the signed-in identity can read SharePoint **and** write to the ADLS file system.

### Two users running simultaneously

Do not share one checkpoint between concurrent processes. Give each process a unique `ADLS_OUTPUT_PREFIX`.

## License

Choose a license appropriate for your organization before publishing publicly.
