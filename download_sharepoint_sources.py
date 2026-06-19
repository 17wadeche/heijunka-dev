#!/usr/bin/env python
from __future__ import annotations
import argparse
import base64
import json
import os
from pathlib import Path
from typing import Any
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
REQUIRED_ENV_VARS = ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET")
def _get_required_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value
def get_access_token() -> str:
    tenant_id = _get_required_env("AZURE_TENANT_ID")
    client_id = _get_required_env("AZURE_CLIENT_ID")
    client_secret = _get_required_env("AZURE_CLIENT_SECRET")
    import msal
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        error = result.get("error") or "unknown_error"
        description = result.get("error_description") or result
        raise RuntimeError(f"Could not acquire Microsoft Graph token: {error}: {description}")
    return result["access_token"]
def encode_share_url(sharepoint_url: str) -> str:
    encoded = base64.urlsafe_b64encode(sharepoint_url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=")
def graph_get(headers: dict[str, str], url: str):
    import requests
    response = requests.get(url, headers=headers, timeout=120)
    response.raise_for_status()
    return response
def download_shared_file(headers: dict[str, str], sharepoint_url: str, local_path: Path) -> None:
    share_id = encode_share_url(sharepoint_url)
    metadata_url = f"{GRAPH_ROOT}/shares/{share_id}/driveItem"
    metadata = graph_get(headers, metadata_url).json()
    if "file" not in metadata:
        raise RuntimeError(f"SharePoint URL did not resolve to a file: {sharepoint_url}")
    download_url = f"{GRAPH_ROOT}/shares/{share_id}/driveItem/content"
    response = graph_get(headers, download_url)
    local_path.parent.mkdir(parents=True, exist_ok=True)
    local_path.write_bytes(response.content)
    print(f"Downloaded {metadata.get('name', local_path.name)} -> {local_path}")
def load_manifest(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as fp:
        manifest = json.load(fp)
    if not isinstance(manifest.get("files"), list):
        raise RuntimeError(f"{path} must contain a 'files' list")
    return manifest
def main() -> None:
    parser = argparse.ArgumentParser(description="Download SharePoint/OneDrive source workbooks for dashboard automation.")
    parser.add_argument("--manifest", default="source_manifest.json", help="Manifest JSON containing files to download.")
    parser.add_argument("--only", action="append", help="Download only entries with this name. Repeatable.")
    args = parser.parse_args()
    selected_names = set(args.only or [])
    manifest = load_manifest(Path(args.manifest))
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    downloaded = 0
    for item in manifest["files"]:
        name = item.get("name")
        if selected_names and name not in selected_names:
            continue
        if item.get("type", "file") != "file":
            raise RuntimeError(f"Only file downloads are currently supported; unsupported entry: {name}")
        sharepoint_url = item.get("sharepoint_url")
        local_path = item.get("local_path")
        if not sharepoint_url or not local_path:
            raise RuntimeError(f"Manifest entry must include sharepoint_url and local_path: {name}")
        download_shared_file(headers, sharepoint_url, Path(local_path))
        downloaded += 1
    if downloaded == 0:
        raise RuntimeError("No manifest entries were downloaded")
if __name__ == "__main__":
    main()