"""Private Google Drive storage. Only folders made by this app are touched."""

from __future__ import annotations

from datetime import datetime, timezone
import io
import re

DRIVE_FILE_SCOPE = "https://www.googleapis.com/auth/drive.file"
FOLDER_MIME = "application/vnd.google-apps.folder"


def get_drive(settings: dict):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    oauth = settings["google_oauth"]
    credentials = Credentials(
        token=None,
        refresh_token=oauth["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=oauth["client_id"],
        client_secret=oauth["client_secret"],
        scopes=[DRIVE_FILE_SCOPE],
    )
    credentials.refresh(Request())
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def escape_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def safe_folder_label(label: str) -> str:
    cleaned = re.sub(r"[\\/\x00-\x1f]", "-", label).strip(" .")
    if not cleaned:
        raise ValueError("Invalid folder name")
    return cleaned[:120]


def find_folder(drive, parent_id: str, name: str) -> str | None:
    query = (
        f"'{escape_drive_query(parent_id)}' in parents and "
        f"name = '{escape_drive_query(name)}' and "
        f"mimeType = '{FOLDER_MIME}' and trashed = false"
    )
    response = drive.files().list(
        q=query, fields="nextPageToken,files(id,name)", pageSize=50,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    matches = response.get("files", [])
    return matches[0]["id"] if matches else None


def ensure_folder(drive, parent_id: str, name: str) -> str:
    name = safe_folder_label(name)
    found = find_folder(drive, parent_id, name)
    if found:
        return found
    created = drive.files().create(
        body={"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]},
        fields="id", supportsAllDrives=True,
    ).execute()
    return created["id"]


def employee_folder_name(employee: dict) -> str:
    return safe_folder_label(employee["name"])


def find_document_folder(drive, root_id: str, employee: dict, document: str) -> str | None:
    emp_id = find_folder(drive, root_id, employee_folder_name(employee))
    if not emp_id:
        return None
    return find_folder(drive, emp_id, safe_folder_label(document))


def has_uploaded(drive, root_id: str, employee: dict, document: str) -> bool:
    folder_id = find_document_folder(drive, root_id, employee, document)
    if not folder_id:
        return False
    response = drive.files().list(
        q=f"'{escape_drive_query(folder_id)}' in parents and trashed = false",
        fields="files(id)", pageSize=1,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return bool(response.get("files"))


FILE_TYPES = {
    ".pdf": ("application/pdf", b"%PDF-"),
    ".jpg": ("image/jpeg", b"\xff\xd8\xff"),
    ".jpeg": ("image/jpeg", b"\xff\xd8\xff"),
    ".png": ("image/png", b"\x89PNG\r\n\x1a\n"),
}


def validate_upload(filename: str, contents: bytes, max_mb: int = 15) -> tuple[str, str]:
    suffix = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if suffix not in FILE_TYPES:
        raise ValueError("Only PDF, JPG, JPEG or PNG files are accepted.")
    if not contents:
        raise ValueError("This file is empty.")
    if len(contents) > max_mb * 1024 * 1024:
        raise ValueError(f"File exceeds the {max_mb} MB limit.")
    mime, magic = FILE_TYPES[suffix]
    if not contents.startswith(magic):
        raise ValueError("File contents do not match its format.")
    return suffix, mime


def upload_document(drive, root_id: str, employee: dict, document: str,
                    original_name: str, contents: bytes) -> str:
    suffix, mime = validate_upload(original_name, contents)
    emp_id = ensure_folder(drive, root_id, employee_folder_name(employee))
    doc_id = ensure_folder(drive, emp_id, document)
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d_%H-%M-%S_%f_UTC")
    safe_doc = safe_folder_label(document)
    filename = f"{safe_doc} - {timestamp}{suffix}"
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(io.BytesIO(contents), mimetype=mime, resumable=True)
    result = drive.files().create(
        body={"name": filename, "parents": [doc_id]},
        media_body=media, fields="id", supportsAllDrives=True,
    ).execute()
    return result["id"]
