"""SharePoint verified-image metadata exporter.

Scans a configured SharePoint folder recursively using Microsoft Graph,
exports only images whose configured verification field is checked, and
captures image/GPS/photo metadata without downloading image files.
"""

# ============================================================
# IMPORTS
# ============================================================

import json
import shutil
import subprocess
import sys
import os

from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse, unquote, quote

import pandas as pd
import requests
import urllib3

from azure.identity import ClientSecretCredential, ManagedIdentityCredential
from azure.storage.filedatalake import DataLakeServiceClient


# ============================================================
# USER SETTINGS
# ============================================================

SITE_URL = os.getenv(
    "SHAREPOINT_SITE_URL",
    "",
).strip()

TENANT_ID = os.getenv(
    "AZURE_TENANT_ID",
    "",
).strip()

# Path INSIDE the SharePoint document library root.
# Do not include "Documents" or "Shared Documents".
FOLDER_PATH_IN_LIBRARY = os.getenv(
    "SHAREPOINT_FOLDER_PATH",
    "",
).strip()

LIBRARY_NAME_CANDIDATES = [
    value.strip()
    for value in os.getenv(
        "SHAREPOINT_LIBRARY_CANDIDATES",
        "Documents,Shared Documents",
    ).split(",")
    if value.strip()
]

# SharePoint column
VERIFIED_BY_UPLOADER_DISPLAY_NAME = os.getenv(
    "VERIFIED_COLUMN_DISPLAY_NAME",
    "Verified By Uploader",
).strip()

# Azure Data Lake Storage Gen2 output configuration.
# All generated artifacts are stored remotely in ADLS.
ADLS_ACCOUNT_URL = os.getenv("ADLS_ACCOUNT_URL", "").strip()
ADLS_FILE_SYSTEM = os.getenv("ADLS_FILE_SYSTEM", "").strip()
ADLS_OUTPUT_PREFIX = os.getenv("ADLS_OUTPUT_PREFIX", "sharepoint-image-metadata-exporter").strip().strip("/")
ADLS_LOCAL_CACHE = os.getenv("ADLS_LOCAL_CACHE", ".cache/output").strip()

# Remote paths inside ADLS.
OUTPUT_EXCEL = os.getenv("OUTPUT_EXCEL", "excel/image_metadata.xlsx").strip().lstrip("/")
OUTPUT_CSV = os.getenv("OUTPUT_CSV", "csv/image_metadata.csv").strip().lstrip("/")
OUTPUT_JSON = os.getenv("OUTPUT_JSON", "json/image_metadata.json").strip().lstrip("/")
DEBUG_CSV = os.getenv("DEBUG_CSV", "debug/debug.csv").strip().lstrip("/")
CHECKPOINT_FILE = os.getenv("CHECKPOINT_FILE", "checkpoint/checkpoint.json").strip().lstrip("/")


# ============================================================
# RESUME OPTIONS
# ============================================================

# False = normal resume behavior.
#
# True = delete existing checkpoint and start over.

RESET_CHECKPOINT = False


# ------------------------------------------------------------
# IMPORTANT
#
# If RESET_CHECKPOINT = True and this is True:
#
#     - old Excel is deleted
#     - old debug CSV is deleted
#     - old checkpoint is deleted
#
# This gives you a completely clean run.
#
# Recommended for the FIRST run of this v5 script.
# ------------------------------------------------------------

CLEAN_OUTPUT_ON_RESET = True


# ============================================================
# OPTIONS
# ============================================================

# Set to False only if your environment requires SSL
# verification to be disabled.
SSL_VERIFY = os.getenv("SSL_VERIFY", "true").strip().lower() not in {"0", "false", "no", "off"}

DEBUG = os.getenv("DEBUG", "false").strip().lower() in {"1", "true", "yes", "on"}

# Save Excel/debug after this many newly processed images.
#
# The checkpoint is STILL saved after EVERY image.

EXCEL_SAVE_INTERVAL = 25


# ============================================================
# IMAGE EXTENSIONS
# ============================================================

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".tif",
    ".tiff",
    ".heic",
    ".heif",
}


# ============================================================
# GRAPH
# ============================================================

GRAPH_ROOT = (
    "https://graph.microsoft.com/v1.0"
)


# ============================================================
# EXCEL COLUMNS
# ============================================================

HEADERS = [
    "FileId",
    "FileName",
    "Folder",
    "Location",
    "Latitude",
    "Longitude",
    "Altitude",
    "ImageTakenDate",
    "FileModifiedDate",
    "FileSizeBytes",
    "ImageWidth",
    "ImageHeight",
    "CameraMake",
    "CameraModel",
    "SourcePath",
    "ImageUrl",
    "VerifiedByUploader",
    "HasGps",
    "LastScanned",
]


# ============================================================
# AUTHENTICATION
# ============================================================

# Azure Function authentication:
#
# 1. User-assigned Managed Identity:
#       AZURE_CLIENT_ID=<managed identity client/application ID>
#       No client secret is required.
#
# 2. Service principal:
#       AZURE_TENANT_ID=<tenant ID>
#       AZURE_CLIENT_ID=<application/client ID>
#       AZURE_CLIENT_SECRET=<client secret>
#
# Azure CLI is intentionally NOT used.

CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "").strip()
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "").strip()


def get_azure_credential():
    """
    Return a credential that works in Azure Functions without Azure CLI.

    If AZURE_CLIENT_SECRET is configured, authenticate as a service
    principal. Otherwise authenticate with the configured user-assigned
    managed identity.
    """
    if CLIENT_SECRET:
        if not TENANT_ID or not CLIENT_ID:
            raise RuntimeError(
                "Service-principal authentication requires "
                "AZURE_TENANT_ID, AZURE_CLIENT_ID and AZURE_CLIENT_SECRET."
            )

        return ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET,
        )

    if CLIENT_ID:
        return ManagedIdentityCredential(
            client_id=CLIENT_ID
        )

    raise RuntimeError(
        "Azure authentication is not configured. For a user-assigned "
        "managed identity set AZURE_CLIENT_ID. For a service principal "
        "set AZURE_TENANT_ID, AZURE_CLIENT_ID and AZURE_CLIENT_SECRET."
    )


AZURE_CREDENTIAL = None


def get_azure_credential_singleton():
    global AZURE_CREDENTIAL

    if AZURE_CREDENTIAL is None:
        AZURE_CREDENTIAL = get_azure_credential()

    return AZURE_CREDENTIAL


def get_graph_token() -> str:
    credential = get_azure_credential_singleton()

    token = credential.get_token(
        "https://graph.microsoft.com/.default"
    )

    if not token or not token.token:
        raise RuntimeError(
            "Azure credential returned an empty Microsoft Graph token."
        )

    return token.token


def make_session() -> requests.Session:
    session = requests.Session()

    session.headers.update(
        {
            "Authorization": f"Bearer {get_graph_token()}",
            "Accept": "application/json",
        }
    )

    session.verify = SSL_VERIFY

    if not SSL_VERIFY:
        urllib3.disable_warnings(
            urllib3.exceptions.InsecureRequestWarning
        )

        print(
            "WARNING: SSL verification is disabled "
            "because SSL_VERIFY is false."
        )

    return session


# ============================================================
# GRAPH HELPERS
# ============================================================

def graph_get(
    session: requests.Session,
    path_or_url: str,
    params=None,
):

    if path_or_url.startswith("https://"):

        url = path_or_url

    else:

        url = (
            f"{GRAPH_ROOT}"
            f"{path_or_url}"
        )

    response = session.get(
        url,
        params=params,
        timeout=60,
    )

    if response.status_code >= 400:

        raise RuntimeError(
            "Graph GET failed.\n"
            f"URL: {response.url}\n"
            f"Status: {response.status_code}\n"
            f"Response: {response.text[:5000]}"
        )

    try:

        return response.json()

    except ValueError as exc:

        raise RuntimeError(
            "Graph returned a non-JSON response.\n"
            f"URL: {response.url}\n"
            f"Status: {response.status_code}\n"
            f"Response: {response.text[:5000]}"
        ) from exc


def graph_get_all(
    session: requests.Session,
    path_or_url: str,
    params=None,
):

    rows = []

    next_url = path_or_url
    next_params = params

    while next_url:

        data = graph_get(
            session,
            next_url,
            params=next_params,
        )

        values = data.get(
            "value",
            []
        )

        if isinstance(values, list):

            rows.extend(
                values
            )

        next_url = data.get(
            "@odata.nextLink"
        )

        next_params = None

    return rows


# ============================================================
# SITE
# ============================================================

def parse_site_url(
    site_url: str,
):

    parsed = urlparse(
        site_url
    )

    if not parsed.netloc:

        raise ValueError(
            f"Invalid SharePoint SITE_URL:\n"
            f"{site_url}"
        )

    site_path = (
        parsed.path.rstrip("/")
    )

    if not site_path:

        raise ValueError(
            "SITE_URL must contain the SharePoint "
            "site path."
        )

    return (
        parsed.netloc,
        site_path,
    )


def get_site(
    session: requests.Session,
):

    hostname, site_path = parse_site_url(
        SITE_URL
    )

    endpoint = (
        f"/sites/{hostname}:{site_path}"
    )

    site = graph_get(
        session,
        endpoint,
    )

    print(
        "Resolved site: "
        f"{site.get('displayName') or site.get('name')}"
    )

    print(
        f"Site ID: {site.get('id')}"
    )

    return site


# ============================================================
# DOCUMENT LIBRARIES / DRIVES
# ============================================================

def get_drives(
    session: requests.Session,
    site_id: str,
):

    return graph_get_all(
        session,
        f"/sites/{site_id}/drives",
        params={
            "$select": (
                "id,name,webUrl,driveType"
            ),
        },
    )


# ============================================================
# SHAREPOINT COLUMNS
# ============================================================

def get_drive_columns(
    session: requests.Session,
    drive_id: str,
):

    return graph_get_all(
        session,
        f"/drives/{drive_id}/list/columns",
        params={
            "$select": (
                "id,name,displayName,"
                "description,columnGroup,"
                "hidden,readOnly"
            ),
        },
    )


def find_verified_by_uploader_column(
    session: requests.Session,
    drive_id: str,
):

    columns = get_drive_columns(
        session,
        drive_id,
    )

    print("")
    print(
        "Searching document-library columns for:"
    )

    print(
        f"  {VERIFIED_BY_UPLOADER_DISPLAY_NAME!r}"
    )

    for column in columns:

        display_name = str(
            column.get(
                "displayName"
            ) or ""
        ).strip()

        internal_name = str(
            column.get(
                "name"
            ) or ""
        ).strip()

        if (
            display_name.lower()
            == VERIFIED_BY_UPLOADER_DISPLAY_NAME.lower()
        ):

            print("")
            print(
                "FOUND Verified By Uploader column:"
            )

            print(
                f"  Display name : "
                f"{display_name}"
            )

            print(
                f"  Internal name: "
                f"{internal_name}"
            )

            print(
                f"  Column ID    : "
                f"{column.get('id')}"
            )

            print("")

            return internal_name

    print("")
    print(
        "Exact column name was not found."
    )

    print("")
    print(
        "Verification/upload-related columns "
        "reported by Graph:"
    )

    found_possible = False

    for column in columns:

        display_name = str(
            column.get(
                "displayName"
            ) or ""
        )

        internal_name = str(
            column.get(
                "name"
            ) or ""
        )

        combined = (
            f"{display_name} "
            f"{internal_name}"
        ).lower()

        if (
            "verif" in combined
            or "upload" in combined
        ):

            found_possible = True

            print(
                f"  displayName={display_name!r}, "
                f"name={internal_name!r}"
            )

    if not found_possible:

        print(
            "  No verification/upload columns found."
        )

    raise RuntimeError(
        "Microsoft Graph could not find the "
        f"SharePoint column "
        f"'{VERIFIED_BY_UPLOADER_DISPLAY_NAME}'.\n\n"
        "The script will not fall back to the old "
        "'Verified' column."
    )


# ============================================================
# PATH HELPERS
# ============================================================

def path_variants(
    path_in_library: str,
):

    base = (
        unquote(path_in_library)
        .replace("\\", "/")
        .strip("/")
    )

    variants = []

    candidates = [
        base,
        base.replace("+", " "),
        base.replace(" ", "+"),
    ]

    for path in candidates:

        path = path.strip("/")

        if path and path not in variants:

            variants.append(
                path
            )

    return variants


def encoded_graph_path(
    path_in_library: str,
):

    clean_path = (
        path_in_library
        .strip("/")
    )

    return "/".join(
        quote(
            segment,
            safe="",
        )
        for segment in clean_path.split("/")
    )


# ============================================================
# FIND TARGET FOLDER
# ============================================================

def get_drive_item_by_path(
    session: requests.Session,
    drive_id: str,
    path_in_library: str,
):

    encoded_path = encoded_graph_path(
        path_in_library
    )

    endpoint = (
        f"/drives/{drive_id}"
        f"/root:/{encoded_path}"
    )

    return graph_get(
        session,
        endpoint,
        params={
            "$expand": (
                "listItem($expand=fields)"
            ),
        },
    )


def find_target_folder(
    session: requests.Session,
    site_id: str,
):

    drives = get_drives(
        session,
        site_id,
    )

    if not drives:

        raise RuntimeError(
            "Microsoft Graph returned no document "
            "libraries/drives for this SharePoint site."
        )

    print("")
    print(
        "Available drives/libraries:"
    )

    for drive in drives:

        print(
            f"  name={drive.get('name')!r}, "
            f"id={drive.get('id')}, "
            f"webUrl={drive.get('webUrl')}"
        )

    print("")

    ordered_drives = []

    for wanted in LIBRARY_NAME_CANDIDATES:

        for drive in drives:

            if (
                str(
                    drive.get(
                        "name",
                        ""
                    )
                ).lower()
                == wanted.lower()
                and drive not in ordered_drives
            ):

                ordered_drives.append(
                    drive
                )

    for drive in drives:

        if drive not in ordered_drives:

            ordered_drives.append(
                drive
            )

    errors = []

    for drive in ordered_drives:

        drive_id = drive["id"]

        drive_name = (
            drive.get("name")
        )

        for folder_path in path_variants(
            FOLDER_PATH_IN_LIBRARY
        ):

            try:

                print(
                    f"Trying drive "
                    f"'{drive_name}' "
                    f"path: "
                    f"{folder_path}"
                )

                folder_item = (
                    get_drive_item_by_path(
                        session,
                        drive_id,
                        folder_path,
                    )
                )

                if "folder" not in folder_item:

                    raise RuntimeError(
                        "Graph found the item, "
                        "but it is not a folder."
                    )

                print(
                    f"Folder found: "
                    f"{folder_item.get('name')}"
                )

                print(
                    f"Folder ID: "
                    f"{folder_item.get('id')}"
                )

                return (
                    drive,
                    folder_path,
                    folder_item,
                )

            except Exception as exc:

                errors.append(
                    f"drive={drive_name}, "
                    f"path={folder_path}, "
                    f"error={exc}"
                )

                if DEBUG:

                    print(
                        f"  Failed: {exc}"
                    )

    raise RuntimeError(
        "Could not find the target folder "
        "using Microsoft Graph.\n\n"
        "Last errors:\n\n"
        + "\n\n".join(
            errors[-10:]
        )
    )


# ============================================================
# RECURSIVE SCANNING
# ============================================================

def get_children_for_item(
    session: requests.Session,
    drive_id: str,
    item_id: str,
):

    endpoint = (
        f"/drives/{drive_id}"
        f"/items/{item_id}/children"
    )

    return graph_get_all(
        session,
        endpoint,
        params={
            "$top": "200",

            "$select": (
                "id,"
                "name,"
                "size,"
                "webUrl,"
                "lastModifiedDateTime,"
                "parentReference,"
                "file,"
                "folder,"
                "listItem"
            ),

            "$expand": (
                "listItem($expand=fields)"
            ),
        },
    )


def recursive_folder_scan(
    session: requests.Session,
    drive_id: str,
    folder_item: dict,
    folder_path: str,
    process_item_callback,
):

    folder_name = (
        folder_item.get(
            "name"
        ) or ""
    )

    folder_item_id = (
        folder_item.get(
            "id"
        )
    )

    if not folder_item_id:

        raise RuntimeError(
            f"Folder '{folder_name}' "
            "has no Graph item ID."
        )

    print("")
    print(
        "------------------------------------------------"
    )

    print(
        f"Scanning folder: "
        f"{folder_path}"
    )

    print(
        "------------------------------------------------"
    )

    children = get_children_for_item(
        session,
        drive_id,
        folder_item_id,
    )

    print(
        f"Children found: "
        f"{len(children)}"
    )

    for item in children:

        item_name = (
            item.get(
                "name"
            ) or ""
        )

        # ------------------------------------------------
        # SUBFOLDER
        # ------------------------------------------------

        if "folder" in item:

            subfolder_path = (
                f"{folder_path}/"
                f"{item_name}"
            )

            print(
                f"  Entering subfolder: "
                f"{item_name}"
            )

            recursive_folder_scan(
                session=session,
                drive_id=drive_id,
                folder_item=item,
                folder_path=subfolder_path,
                process_item_callback=(
                    process_item_callback
                ),
            )

            continue

        # ------------------------------------------------
        # FILE
        # ------------------------------------------------

        if "file" in item:

            process_item_callback(
                item,
                folder_path,
            )


# ============================================================
# FIELD HELPERS
# ============================================================

def get_fields(
    item,
):

    list_item = (
        item.get(
            "listItem"
        ) or {}
    )

    fields = (
        list_item.get(
            "fields"
        ) or {}
    )

    if isinstance(
        fields,
        dict,
    ):

        return fields

    return {}


def get_field(
    fields,
    names,
    default=None,
):

    for name in names:

        if (
            name in fields
            and fields.get(name)
            not in [None, ""]
        ):

            return fields.get(
                name
            )

    return default


# ============================================================
# BOOLEAN
# ============================================================

def truthy_value(
    value,
) -> bool:

    if value is True:

        return True

    if value is False:

        return False

    if value is None:

        return False

    if isinstance(
        value,
        int,
    ):

        return value == 1

    if isinstance(
        value,
        float,
    ):

        return value == 1.0

    if isinstance(
        value,
        str,
    ):

        cleaned = (
            value
            .strip()
            .lower()
        )

        return cleaned in {
            "true",
            "yes",
            "y",
            "1",
            "checked",
            "✓",
            "✔",
        }

    return False


# ============================================================
# JSON
# ============================================================

def safe_json(
    value,
):

    if isinstance(
        value,
        (dict, list),
    ):

        return json.dumps(
            value,
            ensure_ascii=False,
        )

    return value


# ============================================================
# PATH
# ============================================================

def source_path_from_item(
    item,
):

    parent_reference = (
        item.get(
            "parentReference"
        ) or {}
    )

    parent_path = (
        parent_reference.get(
            "path"
        ) or ""
    )

    name = (
        item.get(
            "name"
        ) or ""
    )

    if "root:" in parent_path:

        relative_folder = (
            parent_path
            .split(
                "root:",
                1
            )[1]
            .strip("/")
        )

        if relative_folder:

            return (
                f"/{relative_folder}/"
                f"{name}"
            )

        return f"/{name}"

    return item.get(
        "webUrl"
    )


# ============================================================
# GET COMPLETE IMAGE DRIVE ITEM
#
# THIS IS THE IMPORTANT METADATA FIX.
#
# We explicitly retrieve the individual driveItem so that
# Graph returns:
#
#     location
#     photo
#     image
#
# rather than depending on the /children response.
# ============================================================

def get_complete_image_item(
    session: requests.Session,
    drive_id: str,
    item_id: str,
):

    endpoint = (
        f"/drives/{drive_id}"
        f"/items/{item_id}"
    )

    params = {

        "$select": (
            "id,"
            "name,"
            "size,"
            "webUrl,"
            "lastModifiedDateTime,"
            "parentReference,"
            "file,"
            "folder,"
            "location,"
            "photo,"
            "image,"
            "listItem"
        ),

        "$expand": (
            "listItem($expand=fields)"
        ),
    }

    return graph_get(
        session,
        endpoint,
        params=params,
    )


# ============================================================
# VERIFIED FIELD
# ============================================================

def get_verified_by_uploader_value(
    session: requests.Session,
    drive_id: str,
    item_id: str,
    internal_field_name: str,
):

    endpoint = (
        f"/drives/{drive_id}"
        f"/items/{item_id}"
        f"/listItem/fields"
    )

    data = graph_get(
        session,
        endpoint,
        params={
            "$select": internal_field_name,
        },
    )

    return data.get(
        internal_field_name
    )


# ============================================================
# BUILD ROW
#
# GPS / PHOTO METADATA COMES FROM driveItem.
#
# VERIFIED BY UPLOADER COMES FROM SharePoint fields.
# ============================================================

def build_row(
    item,
    folder_path,
    last_scanned,
    verified_by_uploader,
):

    fields = get_fields(
        item
    )

    # ========================================================
    # GRAPH MEDIA METADATA
    # ========================================================

    location = (
        item.get(
            "location"
        ) or {}
    )

    photo = (
        item.get(
            "photo"
        ) or {}
    )

    image = (
        item.get(
            "image"
        ) or {}
    )

    # ========================================================
    # GPS
    # ========================================================

    latitude = location.get(
        "latitude"
    )

    longitude = location.get(
        "longitude"
    )

    altitude = location.get(
        "altitude"
    )

    # ========================================================
    # PHOTO
    # ========================================================

    image_taken_date = photo.get(
        "takenDateTime"
    )

    camera_make = photo.get(
        "cameraMake"
    )

    camera_model = photo.get(
        "cameraModel"
    )

    # ========================================================
    # IMAGE DIMENSIONS
    # ========================================================

    image_width = image.get(
        "width"
    )

    image_height = image.get(
        "height"
    )

    # ========================================================
    # DEBUG
    # ========================================================

    if DEBUG:

        print(
            "  Graph metadata:"
        )

        print(
            f"    location = "
            f"{location!r}"
        )

        print(
            f"    photo = "
            f"{photo!r}"
        )

        print(
            f"    image = "
            f"{image!r}"
        )

        print(
            f"    Latitude = "
            f"{latitude!r}"
        )

        print(
            f"    Longitude = "
            f"{longitude!r}"
        )

        print(
            f"    Altitude = "
            f"{altitude!r}"
        )

        print(
            f"    ImageTakenDate = "
            f"{image_taken_date!r}"
        )

    # ========================================================
    # RETURN ROW
    # ========================================================

    return {

        "FileId": (
            item.get(
                "id"
            )
            or (
                item.get(
                    "listItem"
                ) or {}
            ).get(
                "id"
            )
        ),

        "FileName": (
            item.get(
                "name"
            )
        ),

        "Folder": (
            folder_path
        ),

        "Location": safe_json(
            location
        ),

        "Latitude": latitude,

        "Longitude": longitude,

        "Altitude": altitude,

        "ImageTakenDate": (
            image_taken_date
        ),

        "FileModifiedDate": (
            item.get(
                "lastModifiedDateTime"
            )
        ),

        "FileSizeBytes": (
            item.get(
                "size"
            )
        ),

        "ImageWidth": (
            image_width
        ),

        "ImageHeight": (
            image_height
        ),

        "CameraMake": (
            camera_make
        ),

        "CameraModel": (
            camera_model
        ),

        "SourcePath": (
            source_path_from_item(
                item
            )
        ),

        "ImageUrl": (
            item.get(
                "webUrl"
            )
        ),

        "VerifiedByUploader": (
            verified_by_uploader
        ),

        "HasGps": (
            latitude is not None
            and longitude is not None
        ),

        "LastScanned": (
            last_scanned
        ),
    }


# ============================================================
# CHECKPOINT
# ============================================================

class AdlsOutputStore:
    """Synchronize local working files with Azure Data Lake Storage Gen2."""

    def __init__(self, account_url, file_system, prefix, local_cache):
        if not account_url or not file_system:
            raise RuntimeError(
                "ADLS_ACCOUNT_URL and ADLS_FILE_SYSTEM are required. "
                "Example: ADLS_ACCOUNT_URL=https://mystorage.dfs.core.windows.net"
            )
        self.account_url = account_url.rstrip("/")
        self.file_system = file_system
        self.prefix = prefix.strip("/")
        self.local_cache = Path(local_cache)
        self.local_cache.mkdir(parents=True, exist_ok=True)
        self.credential = get_azure_credential_singleton()
        self.service = DataLakeServiceClient(
            account_url=self.account_url,
            credential=self.credential,
        )
        self.fs = self.service.get_file_system_client(self.file_system)

    def remote_path(self, relative_path):
        rel = str(relative_path).replace("\\", "/").lstrip("/")
        return f"{self.prefix}/{rel}" if self.prefix else rel

    def local_path(self, relative_path):
        return self.local_cache / Path(str(relative_path))

    def download_if_exists(self, relative_path):
        local = self.local_path(relative_path)
        remote = self.remote_path(relative_path)
        try:
            client = self.fs.get_file_client(remote)
            if not client.exists():
                return False
            local.parent.mkdir(parents=True, exist_ok=True)
            with local.open("wb") as f:
                data = client.download_file()
                f.write(data.readall())
            return True
        except Exception as exc:
            raise RuntimeError(f"Could not download ADLS artifact '{remote}': {exc}") from exc

    def upload(self, relative_path):
        local = self.local_path(relative_path)
        if not local.exists():
            return
        remote = self.remote_path(relative_path)
        try:
            client = self.fs.get_file_client(remote)
            with local.open("rb") as f:
                client.upload_data(f, overwrite=True)
        except Exception as exc:
            raise RuntimeError(f"Could not upload ADLS artifact '{remote}': {exc}") from exc

    def delete(self, relative_path):
        local = self.local_path(relative_path)
        if local.exists():
            local.unlink()
        try:
            client = self.fs.get_file_client(self.remote_path(relative_path))
            if client.exists():
                client.delete_file()
        except Exception as exc:
            raise RuntimeError(f"Could not delete ADLS artifact '{self.remote_path(relative_path)}': {exc}") from exc


OUTPUT_STORE = None


def create_empty_checkpoint(
    drive_id,
    root_folder,
):

    return {

        "version": 2,

        "drive_id": drive_id,

        "root_folder": root_folder,

        "created_at": (
            datetime.now().isoformat(
                timespec="seconds"
            )
        ),

        "updated_at": (
            datetime.now().isoformat(
                timespec="seconds"
            )
        ),

        "processed_file_ids": [],

    }


def load_checkpoint(
    checkpoint_path: Path,
    drive_id: str,
    root_folder: str,
):

    # --------------------------------------------------------
    # RESET
    # --------------------------------------------------------

    if RESET_CHECKPOINT:

        print("")
        print(
            "RESET_CHECKPOINT=True"
        )

        if checkpoint_path.exists():

            print(
                "Deleting existing checkpoint:"
            )

            print(
                checkpoint_path
            )

            checkpoint_path.unlink()

        return create_empty_checkpoint(
            drive_id,
            root_folder,
        )

    # --------------------------------------------------------
    # NO CHECKPOINT
    # --------------------------------------------------------

    if not checkpoint_path.exists():

        print("")
        print(
            "No checkpoint found."
        )

        print(
            "Starting a new scan."
        )

        return create_empty_checkpoint(
            drive_id,
            root_folder,
        )

    # --------------------------------------------------------
    # LOAD
    # --------------------------------------------------------

    print("")
    print(
        "Loading checkpoint:"
    )

    print(
        checkpoint_path
    )

    try:

        with checkpoint_path.open(
            "r",
            encoding="utf-8",
        ) as file:

            checkpoint = json.load(
                file
            )

    except Exception as exc:

        raise RuntimeError(
            "Could not read checkpoint file.\n"
            f"File: {checkpoint_path}\n"
            f"Error: {exc}\n\n"
            "If you are certain the checkpoint is "
            "corrupt, back it up and delete it."
        ) from exc

    # --------------------------------------------------------
    # VALIDATE
    # --------------------------------------------------------

    checkpoint_drive_id = (
        checkpoint.get(
            "drive_id"
        )
    )

    checkpoint_root_folder = (
        checkpoint.get(
            "root_folder"
        )
    )

    if checkpoint_drive_id != drive_id:

        raise RuntimeError(
            "Checkpoint belongs to a different "
            "SharePoint drive/library.\n\n"
            f"Checkpoint drive: "
            f"{checkpoint_drive_id}\n"
            f"Current drive: "
            f"{drive_id}\n\n"
            "Delete or reset the checkpoint if "
            "you intentionally changed libraries."
        )

    if checkpoint_root_folder != root_folder:

        raise RuntimeError(
            "Checkpoint belongs to a different "
            "starting folder.\n\n"
            f"Checkpoint folder: "
            f"{checkpoint_root_folder}\n"
            f"Current folder: "
            f"{root_folder}\n\n"
            "Delete or reset the checkpoint if "
            "you intentionally changed folders."
        )

    if not isinstance(
        checkpoint.get(
            "processed_file_ids"
        ),
        list,
    ):

        raise RuntimeError(
            "Checkpoint has an invalid "
            "'processed_file_ids' value."
        )

    print(
        "Checkpoint loaded successfully."
    )

    print(
        "Previously processed files: "
        f"{len(checkpoint['processed_file_ids'])}"
    )

    return checkpoint


def atomic_save_json(
    path: Path,
    data,
):

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_path = path.with_suffix(
        path.suffix + ".tmp"
    )

    with temp_path.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            data,
            file,
            indent=4,
            ensure_ascii=False,
        )

        file.flush()

        os.fsync(
            file.fileno()
        )

    os.replace(
        temp_path,
        path,
    )

    if OUTPUT_STORE is not None:
        OUTPUT_STORE.upload(path.relative_to(OUTPUT_STORE.local_cache))


def save_checkpoint(
    checkpoint_path: Path,
    checkpoint,
):

    checkpoint["updated_at"] = (
        datetime.now().isoformat(
            timespec="seconds"
        )
    )

    atomic_save_json(
        checkpoint_path,
        checkpoint,
    )


# ============================================================
# EXCEL HELPERS
# ============================================================

def save_excel(
    rows,
    output_path: Path,
):

    if not rows:

        return

    df = pd.DataFrame(
        rows
    )

    for column in HEADERS:

        if column not in df.columns:

            df[column] = None

    df = df[
        HEADERS
    ]

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    df.to_excel(
        output_path,
        index=False,
        sheet_name="ImageMetadata",
        engine="openpyxl",
    )

    if OUTPUT_STORE is not None:
        OUTPUT_STORE.upload(output_path.relative_to(OUTPUT_STORE.local_cache))


def load_existing_excel_rows(
    output_path: Path,
):

    if not output_path.exists():

        return []

    try:

        existing_df = pd.read_excel(
            output_path,
            sheet_name="ImageMetadata",
        )

    except Exception as exc:

        raise RuntimeError(
            "Existing Excel file could not be read.\n"
            f"File: {output_path}\n"
            f"Error: {exc}\n\n"
            "Back up the file and check it before "
            "continuing."
        ) from exc

    if existing_df.empty:

        return []

    existing_df = existing_df.where(
        pd.notna(existing_df),
        None,
    )

    return existing_df.to_dict(
        orient="records"
    )


def save_csv(rows, csv_path: Path):
    if not rows:
        return
    df = pd.DataFrame(rows)
    for column in HEADERS:
        if column not in df.columns:
            df[column] = None
    df = df[HEADERS]
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    if OUTPUT_STORE is not None:
        OUTPUT_STORE.upload(csv_path.relative_to(OUTPUT_STORE.local_cache))


def save_json(rows, json_path: Path):
    if not rows:
        return
    json_path.parent.mkdir(parents=True, exist_ok=True)
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2, ensure_ascii=False, default=str)
        f.flush()
        os.fsync(f.fileno())
    if OUTPUT_STORE is not None:
        OUTPUT_STORE.upload(json_path.relative_to(OUTPUT_STORE.local_cache))


# ============================================================
# DEBUG CSV HELPERS
# ============================================================

def save_debug_csv(
    debug_rows,
    debug_path: Path,
):

    debug_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if not debug_rows:

        return

    pd.DataFrame(
        debug_rows
    ).to_csv(
        debug_path,
        index=False,
        encoding="utf-8-sig",
    )

    if OUTPUT_STORE is not None:
        OUTPUT_STORE.upload(debug_path.relative_to(OUTPUT_STORE.local_cache))


def load_existing_debug_rows(
    debug_path: Path,
):

    if not debug_path.exists():

        return []

    try:

        debug_df = pd.read_csv(
            debug_path
        )

        if debug_df.empty:

            return []

        debug_df = debug_df.where(
            pd.notna(debug_df),
            None,
        )

        return debug_df.to_dict(
            orient="records"
        )

    except Exception:

        print(
            "WARNING: Existing debug CSV "
            "could not be loaded."
        )

        print(
            "A new debug CSV will be created."
        )

        return []


# ============================================================
# CLEAN RESET
# ============================================================

def clean_outputs_for_reset(
    output_path: Path,
    csv_path: Path,
    json_path: Path,
    debug_path: Path,
    checkpoint_path: Path,
):

    if not RESET_CHECKPOINT:

        return

    if not CLEAN_OUTPUT_ON_RESET:

        return

    print("")
    print(
        "CLEAN_OUTPUT_ON_RESET=True"
    )

    paths_to_remove = [
        output_path,
        csv_path,
        json_path,
        debug_path,
        checkpoint_path,
    ]

    for path in paths_to_remove:

        if path.exists():

            print(
                f"Deleting: {path}"
            )

            path.unlink()

    if OUTPUT_STORE is not None:
        for relative_path in [
            output_path.relative_to(OUTPUT_STORE.local_cache),
            csv_path.relative_to(OUTPUT_STORE.local_cache),
            json_path.relative_to(OUTPUT_STORE.local_cache),
            debug_path.relative_to(OUTPUT_STORE.local_cache),
            checkpoint_path.relative_to(OUTPUT_STORE.local_cache),
        ]:
            OUTPUT_STORE.delete(relative_path)


# ============================================================
# MAIN
# ============================================================

def validate_configuration():
    """Validate required configuration before contacting Microsoft Graph."""
    missing = []
    if not SITE_URL:
        missing.append("SHAREPOINT_SITE_URL")
    if not TENANT_ID and CLIENT_SECRET:
        missing.append("AZURE_TENANT_ID")

    if not CLIENT_ID:
        missing.append("AZURE_CLIENT_ID")

    if CLIENT_SECRET and not TENANT_ID:
        missing.append("AZURE_TENANT_ID")
    if not FOLDER_PATH_IN_LIBRARY:
        missing.append("SHAREPOINT_FOLDER_PATH")
    if not ADLS_ACCOUNT_URL:
        missing.append("ADLS_ACCOUNT_URL")
    if not ADLS_FILE_SYSTEM:
        missing.append("ADLS_FILE_SYSTEM")

    if missing:
        raise RuntimeError(
            "Missing required configuration: " + ", ".join(missing) + "\\n"
            "Copy .env.example to .env and fill in the values."
        )


def main():
    validate_configuration()

    print("")
    print(
        "================================================"
    )

    print(
        "Verified By Uploader"
    )

    print(
        "Recursive Graph Image Metadata Export"
    )

    print(
        "V5 - GPS / PHOTO METADATA FIX"
    )

    print(
        "WITH CHECKPOINT / RESUME"
    )

    print(
        "================================================"
    )

    print("")
    print(
        "Metadata-only mode."
    )

    print(
        "No image files will be downloaded."
    )

    print(
        "Only the configured folder and "
        "its subfolders will be scanned."
    )

    print(
        "Only images with "
        "'Verified By Uploader' checked "
        "will be exported."
    )

    print("")
    print(
        "Graph metadata source:"
    )

    print(
        "  GPS:"
    )

    print(
        "    driveItem.location"
    )

    print(
        "  Photo date:"
    )

    print(
        "    driveItem.photo.takenDateTime"
    )

    print(
        "  Image dimensions:"
    )

    print(
        "    driveItem.image"
    )

    print("")

    # --------------------------------------------------------
    # OUTPUT PATHS
    # --------------------------------------------------------

    global OUTPUT_STORE
    OUTPUT_STORE = AdlsOutputStore(
        ADLS_ACCOUNT_URL,
        ADLS_FILE_SYSTEM,
        ADLS_OUTPUT_PREFIX,
        ADLS_LOCAL_CACHE,
    )

    output_path = OUTPUT_STORE.local_path(OUTPUT_EXCEL)
    csv_path = OUTPUT_STORE.local_path(OUTPUT_CSV)
    json_path = OUTPUT_STORE.local_path(OUTPUT_JSON)
    debug_path = OUTPUT_STORE.local_path(DEBUG_CSV)
    checkpoint_path = OUTPUT_STORE.local_path(CHECKPOINT_FILE)

    for remote_path in [OUTPUT_EXCEL, OUTPUT_CSV, OUTPUT_JSON, DEBUG_CSV, CHECKPOINT_FILE]:
        OUTPUT_STORE.download_if_exists(remote_path)

    # --------------------------------------------------------
    # CLEAN RESET
    # --------------------------------------------------------

    clean_outputs_for_reset(
        output_path,
        csv_path,
        json_path,
        debug_path,
        checkpoint_path,
    )

    # --------------------------------------------------------
    # AUTH
    # --------------------------------------------------------

    print(
        "Getting Microsoft Graph token..."
    )

    session = make_session()

    print(
        "Graph authentication successful."
    )

    # --------------------------------------------------------
    # SITE
    # --------------------------------------------------------

    site = get_site(
        session
    )

    site_id = site["id"]

    # --------------------------------------------------------
    # TARGET FOLDER
    # --------------------------------------------------------

    (
        drive,
        folder_path,
        folder_item,
    ) = find_target_folder(
        session,
        site_id,
    )

    drive_id = drive["id"]

    print("")
    print(
        "Starting folder:"
    )

    print(
        f"  Library: "
        f"{drive.get('name')}"
    )

    print(
        f"  Folder: "
        f"{folder_path}"
    )

    print(
        f"  Folder ID: "
        f"{folder_item.get('id')}"
    )

    # --------------------------------------------------------
    # VERIFIED COLUMN
    # --------------------------------------------------------

    verified_column_name = (
        find_verified_by_uploader_column(
            session,
            drive_id,
        )
    )

    print(
        "Using ONLY this internal field:"
    )

    print(
        f"  {verified_column_name}"
    )

    print("")

    # --------------------------------------------------------
    # CHECKPOINT
    # --------------------------------------------------------

    checkpoint = load_checkpoint(
        checkpoint_path,
        drive_id,
        folder_path,
    )

    processed_file_ids = set(
        checkpoint.get(
            "processed_file_ids",
            [],
        )
    )

    # --------------------------------------------------------
    # EXISTING EXCEL
    # --------------------------------------------------------

    rows = load_existing_excel_rows(
        output_path
    )

    if rows:

        print(
            "Existing Excel rows loaded: "
            f"{len(rows)}"
        )

    else:

        print(
            "No existing Excel rows found."
        )

    # --------------------------------------------------------
    # DEBUG
    # --------------------------------------------------------

    debug_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    debug_rows = load_existing_debug_rows(
        debug_path
    )

    if debug_rows:

        print(
            "Existing debug rows loaded: "
            f"{len(debug_rows)}"
        )

    # --------------------------------------------------------
    # COUNTERS
    # --------------------------------------------------------

    last_scanned = (
        datetime.now()
        .isoformat(
            timespec="seconds"
        )
    )

    scanned_files = 0

    skipped_files = 0

    image_count = 0

    verified_count = 0

    verification_errors = 0

    metadata_errors = 0

    newly_processed = 0

    gps_count = 0

    taken_date_count = 0

    # --------------------------------------------------------
    # PROCESS ONE FILE
    # --------------------------------------------------------

    def process_item(
        item,
        current_folder_path,
    ):

        nonlocal scanned_files
        nonlocal skipped_files
        nonlocal image_count
        nonlocal verified_count
        nonlocal verification_errors
        nonlocal metadata_errors
        nonlocal newly_processed
        nonlocal gps_count
        nonlocal taken_date_count

        scanned_files += 1

        item_id = item.get(
            "id"
        )

        name = (
            item.get(
                "name"
            ) or ""
        )

        extension = (
            Path(name)
            .suffix
            .lower()
        )

        is_file = (
            "file" in item
        )

        is_image = (
            is_file
            and extension
            in IMAGE_EXTENSIONS
        )

        # ----------------------------------------------------
        # Ignore non-images
        # ----------------------------------------------------

        if not is_image:

            return

        image_count += 1

        # ----------------------------------------------------
        # Validate FileId
        # ----------------------------------------------------

        if not item_id:

            print("")
            print(
                "WARNING: Image has no FileId."
            )

            print(
                f"  File: {name}"
            )

            return

        # ----------------------------------------------------
        # CHECKPOINT
        # ----------------------------------------------------

        if item_id in processed_file_ids:

            skipped_files += 1

            if DEBUG:

                print(
                    f"RESUME SKIP: "
                    f"{name}"
                )

            return

        # ----------------------------------------------------
        # PROCESS NEW IMAGE
        # ----------------------------------------------------

        print("")
        print(
            "================================================"
        )

        print(
            "PROCESSING IMAGE"
        )

        print(
            "================================================"
        )

        print(
            f"  File: {name}"
        )

        print(
            f"  Folder: "
            f"{current_folder_path}"
        )

        print(
            f"  FileId: "
            f"{item_id}"
        )

        # ----------------------------------------------------
        # GET COMPLETE DRIVE ITEM
        #
        # This is the critical metadata request.
        # ----------------------------------------------------

        metadata_error = None

        try:

            complete_item = (
                get_complete_image_item(
                    session,
                    drive_id,
                    item_id,
                )
            )

            # Use the complete response from this point
            # forward.

            item_for_processing = (
                complete_item
            )

        except Exception as exc:

            metadata_errors += 1

            metadata_error = str(
                exc
            )

            print("")
            print(
                "  ERROR reading complete "
                "Graph image metadata:"
            )

            print(
                f"  {metadata_error}"
            )

            # ------------------------------------------------
            # DO NOT checkpoint this image.
            #
            # That way the next run will retry it.
            # ------------------------------------------------

            return

        # ----------------------------------------------------
        # GRAPH MEDIA METADATA
        # ----------------------------------------------------

        location = (
            item_for_processing.get(
                "location"
            ) or {}
        )

        photo = (
            item_for_processing.get(
                "photo"
            ) or {}
        )

        image = (
            item_for_processing.get(
                "image"
            ) or {}
        )

        latitude = location.get(
            "latitude"
        )

        longitude = location.get(
            "longitude"
        )

        altitude = location.get(
            "altitude"
        )

        image_taken_date = photo.get(
            "takenDateTime"
        )

        image_width = image.get(
            "width"
        )

        image_height = image.get(
            "height"
        )

        camera_make = photo.get(
            "cameraMake"
        )

        camera_model = photo.get(
            "cameraModel"
        )

        if (
            latitude is not None
            and longitude is not None
        ):

            gps_count += 1

        if image_taken_date not in [
            None,
            "",
        ]:

            taken_date_count += 1

        # ----------------------------------------------------
        # PRINT GRAPH METADATA
        # ----------------------------------------------------

        print("")
        print(
            "  GRAPH MEDIA METADATA:"
        )

        print(
            f"    location: "
            f"{location!r}"
        )

        print(
            f"    photo: "
            f"{photo!r}"
        )

        print(
            f"    image: "
            f"{image!r}"
        )

        print(
            f"    Latitude: "
            f"{latitude!r}"
        )

        print(
            f"    Longitude: "
            f"{longitude!r}"
        )

        print(
            f"    Altitude: "
            f"{altitude!r}"
        )

        print(
            f"    ImageTakenDate: "
            f"{image_taken_date!r}"
        )

        print(
            f"    ImageWidth: "
            f"{image_width!r}"
        )

        print(
            f"    ImageHeight: "
            f"{image_height!r}"
        )

        print(
            f"    CameraMake: "
            f"{camera_make!r}"
        )

        print(
            f"    CameraModel: "
            f"{camera_model!r}"
        )

        # ----------------------------------------------------
        # FIELDS
        # ----------------------------------------------------

        fields = get_fields(
            item_for_processing
        )

        # ----------------------------------------------------
        # VERIFIED BY UPLOADER
        # ----------------------------------------------------

        verified_by_uploader = None

        verification_error = None

        try:

            verified_by_uploader = (
                get_verified_by_uploader_value(
                    session,
                    drive_id,
                    item_id,
                    verified_column_name,
                )
            )

        except Exception as exc:

            verification_errors += 1

            verification_error = str(
                exc
            )

            print("")
            print(
                "  ERROR reading "
                "Verified By Uploader:"
            )

            print(
                f"  {verification_error}"
            )

        is_verified = truthy_value(
            verified_by_uploader
        )

        print(
            f"  Verified By Uploader: "
            f"{verified_by_uploader!r}"
        )

        print(
            f"  Interpreted as checked: "
            f"{is_verified}"
        )

        # ----------------------------------------------------
        # DEBUG ROW
        # ----------------------------------------------------

        debug_rows.append(
            {
                "Name": name,

                "Extension": extension,

                "IsFile": is_file,

                "IsImage": is_image,

                "Folder": (
                    current_folder_path
                ),

                "FileId": item_id,

                "VerifiedByUploader": (
                    verified_by_uploader
                ),

                "VerifiedByUploaderInterpreted": (
                    is_verified
                ),

                "VerificationReadError": (
                    verification_error
                ),

                "MetadataReadError": (
                    metadata_error
                ),

                "WebUrl": (
                    item_for_processing.get(
                        "webUrl"
                    )
                ),

                "SourcePath": (
                    source_path_from_item(
                        item_for_processing
                    )
                ),

                # --------------------------------------------
                # RAW GRAPH MEDIA METADATA
                # --------------------------------------------

                "GraphLocation": safe_json(
                    location
                ),

                "GraphPhoto": safe_json(
                    photo
                ),

                "GraphImage": safe_json(
                    image
                ),

                "GraphLatitude": latitude,

                "GraphLongitude": longitude,

                "GraphAltitude": altitude,

                "GraphTakenDateTime": (
                    image_taken_date
                ),

                "GraphImageWidth": (
                    image_width
                ),

                "GraphImageHeight": (
                    image_height
                ),

                "GraphCameraMake": (
                    camera_make
                ),

                "GraphCameraModel": (
                    camera_model
                ),

                # --------------------------------------------
                # SHAREPOINT FIELD KEYS
                # --------------------------------------------

                "AllFieldKeys": (
                    "; ".join(
                        sorted(
                            str(k)
                            for k in fields.keys()
                        )
                    )
                ),
            }
        )

        # ----------------------------------------------------
        # ADD VERIFIED IMAGE TO EXCEL DATA
        # ----------------------------------------------------

        if is_verified:

            verified_count += 1

            row = build_row(
                item_for_processing,
                current_folder_path,
                last_scanned,
                verified_by_uploader,
            )

            rows.append(
                row
            )

            print("")
            print(
                "  >>> VERIFIED IMAGE ADDED TO EXPORT"
            )

            print(
                f"      Latitude: "
                f"{latitude!r}"
            )

            print(
                f"      Longitude: "
                f"{longitude!r}"
            )

            print(
                f"      ImageTakenDate: "
                f"{image_taken_date!r}"
            )

        else:

            print(
                "  Image NOT exported because "
                "Verified By Uploader is not checked."
            )

        # ----------------------------------------------------
        # IMPORTANT
        #
        # Mark FileId processed AFTER:
        #
        # 1. Complete metadata read
        # 2. Verification read
        # 3. Export row decision
        #
        # If metadata retrieval fails, the image is NOT
        # checkpointed and will be retried.
        # ----------------------------------------------------

        processed_file_ids.add(
            item_id
        )

        checkpoint[
            "processed_file_ids"
        ] = list(
            processed_file_ids
        )

        # ----------------------------------------------------
        # SAVE CHECKPOINT IMMEDIATELY
        # ----------------------------------------------------

        save_checkpoint(
            checkpoint_path,
            checkpoint,
        )

        newly_processed += 1

        print("")
        print(
            "  Checkpoint saved."
        )

        print(
            f"  Total checkpointed: "
            f"{len(processed_file_ids)}"
        )

        # ----------------------------------------------------
        # SAVE EXCEL PERIODICALLY
        # ----------------------------------------------------

        if (
            newly_processed
            % EXCEL_SAVE_INTERVAL
            == 0
        ):

            print("")
            print(
                "Saving Excel progress..."
            )

            save_excel(rows, output_path)
            save_csv(rows, csv_path)
            save_json(rows, json_path)

            print(
                f"  Excel rows saved: "
                f"{len(rows)}"
            )

            print(
                f"  Checkpointed files: "
                f"{len(processed_file_ids)}"
            )

            print(
                "Saving debug CSV..."
            )

            save_debug_csv(
                debug_rows,
                debug_path,
            )

    # --------------------------------------------------------
    # START RECURSIVE SCAN
    # --------------------------------------------------------

    print("")
    print(
        "================================================"
    )

    print(
        "STARTING / RESUMING RECURSIVE SCAN"
    )

    print(
        "================================================"
    )

    print(
        f"Already processed: "
        f"{len(processed_file_ids)}"
    )

    print("")

    try:

        recursive_folder_scan(
            session=session,
            drive_id=drive_id,
            folder_item=folder_item,
            folder_path=folder_path,
            process_item_callback=process_item,
        )

    except KeyboardInterrupt:

        print("")
        print(
            "================================================"
        )

        print(
            "SCAN INTERRUPTED"
        )

        print(
            "================================================"
        )

        print(
            "Saving current Excel progress..."
        )

        save_excel(rows, output_path)
        save_csv(rows, csv_path)
        save_json(rows, json_path)

        print(
            "Saving current debug CSV..."
        )

        save_debug_csv(
            debug_rows,
            debug_path,
        )

        print(
            "Checkpoint is already saved "
            "after each successfully processed image."
        )

        print("")
        print(
            "Run the script again to resume."
        )

        raise

    # --------------------------------------------------------
    # FINAL SAVE
    # --------------------------------------------------------

    print("")
    print(
        "Final Excel save..."
    )

    save_excel(rows, output_path)
    save_csv(rows, csv_path)
    save_json(rows, json_path)

    print(
        "Final debug CSV save..."
    )

    save_debug_csv(
        debug_rows,
        debug_path,
    )

    # --------------------------------------------------------
    # FINAL CHECKPOINT
    # --------------------------------------------------------

    checkpoint[
        "completed"
    ] = True

    checkpoint[
        "completed_at"
    ] = (
        datetime.now().isoformat(
            timespec="seconds"
        )
    )

    checkpoint[
        "processed_file_ids"
    ] = list(
        processed_file_ids
    )

    save_checkpoint(
        checkpoint_path,
        checkpoint,
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    print("")
    print(
        "================================================"
    )

    print(
        "FINAL SUMMARY"
    )

    print(
        "================================================"
    )

    print(
        f"Library: "
        f"{drive.get('name')}"
    )

    print(
        f"Starting folder: "
        f"{folder_path}"
    )

    print(
        f"Verified field: "
        f"{verified_column_name}"
    )

    print("")
    print(
        "FILES:"
    )

    print(
        f"  Files encountered: "
        f"{scanned_files}"
    )

    print(
        f"  Images encountered: "
        f"{image_count}"
    )

    print(
        f"  Images skipped from checkpoint: "
        f"{skipped_files}"
    )

    print(
        f"  New images processed this run: "
        f"{newly_processed}"
    )

    print("")
    print(
        "VERIFICATION:"
    )

    print(
        f"  Verified images exported: "
        f"{verified_count}"
    )

    print(
        f"  Verification errors: "
        f"{verification_errors}"
    )

    print("")
    print(
        "GRAPH MEDIA METADATA:"
    )

    print(
        f"  Images with GPS: "
        f"{gps_count}"
    )

    print(
        f"  Images with ImageTakenDate: "
        f"{taken_date_count}"
    )

    print(
        f"  Metadata retrieval errors: "
        f"{metadata_errors}"
    )

    print("")
    print(
        "CHECKPOINT:"
    )

    print(
        f"  Total checkpointed files: "
        f"{len(processed_file_ids)}"
    )

    print("")
    print(
        "Excel:"
    )

    print(
        f"  {output_path} (synced to ADLS)"
    )

    print("")
    print(
        "Debug CSV:"
    )

    print(
        f"  {debug_path} (synced to ADLS)"
    )

    print("")
    print(
        "Checkpoint:"
    )

    print(
        f"  {checkpoint_path}"
    )

    print("")
    print(
        "================================================"
    )

    print(
        "DONE"
    )

    print(
        "================================================"
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except subprocess.CalledProcessError as exc:

        print("")
        print(
            "ERROR running Azure CLI command."
        )

        print(
            f"Command: {exc.cmd}"
        )

        print(
            f"Return code: {exc.returncode}"
        )

        print(
            f"STDOUT: {exc.stdout}"
        )

        print(
            f"STDERR: {exc.stderr}"
        )

        sys.exit(1)

    except KeyboardInterrupt:

        print("")
        print(
            "Stopped."
        )

        print(
            "The checkpoint has been saved."
        )

        print(
            "Run the script again to resume."
        )

        sys.exit(1)

    except Exception as exc:

        print("")
        print(
            "ERROR:"
        )

        print(
            exc
        )

        print("")
        print(
            "If this happened during scanning, "
            "successfully processed images remain "
            "in the checkpoint."
        )

        print(
            "Images that failed during metadata retrieval "
            "were not checkpointed and will be retried."
        )

        print(
            "Run the script again to resume."
        )

        sys.exit(1)

