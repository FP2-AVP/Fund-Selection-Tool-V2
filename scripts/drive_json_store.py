#!/usr/bin/env python3
"""Small Google Drive JSON storage helper for local server APIs."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any


SCOPES = ["https://www.googleapis.com/auth/drive"]


def credentials_from_env() -> Any:
    from google.oauth2 import service_account

    raw_json = (
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_UPLOAD", "").strip()
        or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_EXPORT", "").strip()
        or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    )
    if raw_json:
        info = json.loads(raw_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if credentials_path:
        return service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=SCOPES,
        )

    raise RuntimeError(
        "Set GOOGLE_SERVICE_ACCOUNT_JSON_UPLOAD, GOOGLE_SERVICE_ACCOUNT_JSON, "
        "or GOOGLE_APPLICATION_CREDENTIALS first."
    )


def quote_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def drive_client() -> Any:
    from googleapiclient.discovery import build

    return build("drive", "v3", credentials=credentials_from_env(), cache_discovery=False)


def find_file_in_folder(drive: Any, folder_id: str, name: str) -> str | None:
    response = (
        drive.files()
        .list(
            q=(
                f"name='{quote_drive_query(name)}' "
                f"and '{quote_drive_query(folder_id)}' in parents "
                "and trashed=false"
            ),
            fields="files(id,name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    return files[0]["id"] if files else None


def find_child_folder(drive: Any, parent_id: str, name: str) -> str | None:
    response = (
        drive.files()
        .list(
            q=(
                "mimeType='application/vnd.google-apps.folder' "
                f"and name='{quote_drive_query(name)}' "
                f"and '{quote_drive_query(parent_id)}' in parents "
                "and trashed=false"
            ),
            fields="files(id,name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    return files[0]["id"] if files else None


def create_child_folder(drive: Any, parent_id: str, name: str) -> str:
    result = (
        drive.files()
        .create(
            body={
                "name": name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id],
            },
            fields="id",
            supportsAllDrives=True,
        )
        .execute()
    )
    return result["id"]


def ensure_child_folder(drive: Any, parent_id: str, name: str) -> str:
    return find_child_folder(drive, parent_id, name) or create_child_folder(drive, parent_id, name)


def resolve_folder_path(drive: Any, root_folder_id: str, path_segments: list[str]) -> str:
    folder_id = root_folder_id
    for segment in path_segments:
        clean_segment = str(segment or "").strip()
        if not clean_segment:
            continue
        folder_id = ensure_child_folder(drive, folder_id, clean_segment)
    return folder_id


def find_folder_path(drive: Any, root_folder_id: str, path_segments: list[str]) -> str | None:
    """Resolve an existing folder path without creating missing folders."""
    folder_id = root_folder_id
    for segment in path_segments:
        clean_segment = str(segment or "").strip()
        if not clean_segment:
            continue
        folder_id = find_child_folder(drive, folder_id, clean_segment)
        if not folder_id:
            return None
    return folder_id


def upload_json_payload(folder_id: str, file_name: str, payload: Any) -> str:
    from googleapiclient.http import MediaFileUpload

    drive = drive_client()

    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
        tmp_path = Path(fh.name)

    try:
        media = MediaFileUpload(str(tmp_path), mimetype="application/json", resumable=False)
        existing_id = find_file_in_folder(drive, folder_id, file_name)
        metadata = {
            "name": file_name,
            "mimeType": "application/json",
        }
        if existing_id:
            result = (
                drive.files()
                .update(
                    fileId=existing_id,
                    body=metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )
        else:
            metadata["parents"] = [folder_id]
            result = (
                drive.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )
        return result["id"]
    finally:
        tmp_path.unlink(missing_ok=True)


def upload_json_payload_to_path(
    root_folder_id: str,
    path_segments: list[str],
    file_name: str,
    payload: Any,
) -> tuple[str, str]:
    drive = drive_client()
    folder_id = resolve_folder_path(drive, root_folder_id, path_segments)

    from googleapiclient.http import MediaFileUpload

    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
        tmp_path = Path(fh.name)

    try:
        media = MediaFileUpload(str(tmp_path), mimetype="application/json", resumable=False)
        existing_id = find_file_in_folder(drive, folder_id, file_name)
        metadata = {
            "name": file_name,
            "mimeType": "application/json",
        }
        if existing_id:
            result = (
                drive.files()
                .update(
                    fileId=existing_id,
                    body=metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )
        else:
            metadata["parents"] = [folder_id]
            result = (
                drive.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )
        return result["id"], folder_id
    finally:
        tmp_path.unlink(missing_ok=True)


def download_json_payload(folder_id: str, file_name: str) -> Any | None:
    drive = drive_client()
    file_id = find_file_in_folder(drive, folder_id, file_name)
    if not file_id:
        return None
    response = drive.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
    if isinstance(response, bytes):
        text = response.decode("utf-8")
    else:
        text = str(response)
    return json.loads(text)


def download_json_payload_from_path(
    root_folder_id: str,
    path_segments: list[str],
    file_name: str,
) -> tuple[Any | None, str | None, str | None]:
    """Read JSON from an existing Drive path without creating folders."""
    drive = drive_client()
    folder_id = find_folder_path(drive, root_folder_id, path_segments)
    if not folder_id:
        return None, None, None
    file_id = find_file_in_folder(drive, folder_id, file_name)
    if not file_id:
        return None, folder_id, None
    response = drive.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
    text = response.decode("utf-8") if isinstance(response, bytes) else str(response)
    return json.loads(text), folder_id, file_id


def delete_json_file(folder_id: str, file_name: str) -> bool:
    drive = drive_client()
    file_id = find_file_in_folder(drive, folder_id, file_name)
    if not file_id:
        return False
    drive.files().delete(fileId=file_id, supportsAllDrives=True).execute()
    return True
