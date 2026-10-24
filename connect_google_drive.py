"""Local one-time owner authorisation for a Google Drive upload folder."""

from __future__ import annotations

import argparse
from pathlib import Path
import json
import os

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from drive_storage import DRIVE_FILE_SCOPE, FOLDER_MIME


def run(client_secret_file: Path, output_file: Path, folder_name: str):
    if output_file.exists():
        raise FileExistsError("Google secrets file already exists. Do not overwrite it accidentally.")
    flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_file), [DRIVE_FILE_SCOPE])
    credentials = flow.run_local_server(port=0, prompt="consent", access_type="offline")
    if not credentials.refresh_token:
        raise RuntimeError("No refresh token received. Revoke previous app access and retry authorisation.")
    drive = build("drive", "v3", credentials=credentials, cache_discovery=False)
    created = drive.files().create(
        body={"name": folder_name, "mimeType": FOLDER_MIME},
        fields="id,webViewLink", supportsAllDrives=True,
    ).execute()
    # JSON double-quoted strings are valid TOML basic strings for these ASCII credentials.
    content = (
        f"root_folder_id = {json.dumps(created['id'])}\n\n"
        "[google_oauth]\n"
        f"client_id = {json.dumps(credentials.client_id)}\n"
        f"client_secret = {json.dumps(credentials.client_secret)}\n"
        f"refresh_token = {json.dumps(credentials.refresh_token)}\n"
    )
    output_file.write_text(content, encoding="utf-8")
    os.chmod(output_file, 0o600)
    print(f"Private credentials saved to {output_file}")
    print(f"Google Drive folder ID: {created['id']}")
    print("Open: https://drive.google.com/drive/folders/" + created["id"])
    print("Never commit or share the generated TOML file.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--client-secret", type=Path, default=Path("client_secret.json"))
    parser.add_argument("--output", type=Path, default=Path("google_secrets.toml"))
    parser.add_argument("--folder-name", default="Employee Documents")
    args = parser.parse_args()
    run(args.client_secret, args.output, args.folder_name)
