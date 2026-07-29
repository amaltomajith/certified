import os
import json
import logging
import threading
from pathlib import Path
from typing import Dict, Any, Optional, List

logger = logging.getLogger(__name__)

# Root directory of project
ROOT = Path(__file__).resolve().parent.parent

DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']

# OAuth2 token storage (persisted between runs)
OAUTH_TOKEN_FILE = ROOT / ".drive_oauth_token.json"
# OAuth2 client secrets (uploaded by user)
OAUTH_CLIENT_SECRETS_FILE = ROOT / "oauth_client_secrets.json"

# In-memory OAuth2 flow state (used during the auth flow)
_oauth_flow_state: dict = {"status": "idle", "error": None}  # status: idle|in_progress|done|error


# ─── OAuth2 helpers ────────────────────────────────────────────────────────────

def is_oauth_client_available() -> bool:
    """Return True if the user has uploaded their OAuth2 client_secrets.json."""
    return OAUTH_CLIENT_SECRETS_FILE.exists()


def is_oauth_token_available() -> bool:
    """Return True if a persisted OAuth2 token exists."""
    return OAUTH_TOKEN_FILE.exists()


def is_oauth_drive_available() -> bool:
    """Return True if OAuth2 flow is fully set up (client secrets + token)."""
    try:
        import googleapiclient.discovery
        from google.oauth2.credentials import Credentials
    except ImportError:
        return False
    return is_oauth_client_available() and is_oauth_token_available()


def get_oauth_drive_service():
    """Build and return an authenticated Google Drive service using user OAuth2 credentials."""
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    token_data = json.loads(OAUTH_TOKEN_FILE.read_text(encoding="utf-8"))
    creds = Credentials(
        token=token_data.get("token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=token_data.get("token_uri"),
        client_id=token_data.get("client_id"),
        client_secret=token_data.get("client_secret"),
        scopes=token_data.get("scopes"),
    )

    # Refresh if expired
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        # Save refreshed token
        updated = json.loads(OAUTH_TOKEN_FILE.read_text(encoding="utf-8"))
        updated["token"] = creds.token
        OAUTH_TOKEN_FILE.write_text(json.dumps(updated), encoding="utf-8")

    return build("drive", "v3", credentials=creds)


def get_oauth_user_email() -> Optional[str]:
    """Return the email of the authorized Google account from saved token."""
    if not OAUTH_TOKEN_FILE.exists():
        return None
    try:
        data = json.loads(OAUTH_TOKEN_FILE.read_text(encoding="utf-8"))
        return data.get("user_email")
    except Exception:
        return None


_oauth_flow_lock = threading.Lock()


def reset_oauth_flow_state() -> None:
    """Reset the in-memory OAuth flow state."""
    with _oauth_flow_lock:
        _oauth_flow_state["status"] = "idle"
        _oauth_flow_state["error"] = None


def start_oauth_flow_in_background(force: bool = True) -> None:
    """Start the OAuth2 InstalledAppFlow in a background thread (opens browser)."""
    with _oauth_flow_lock:
        if not force and _oauth_flow_state.get("status") == "in_progress":
            return
        _oauth_flow_state["status"] = "in_progress"
        _oauth_flow_state["error"] = None

    def _run():
        import sys
        print("[OAuth2] Starting Google Drive authorization flow in background thread...", flush=True)
        try:
            try:
                from google_auth_oauthlib.flow import InstalledAppFlow, WSGITimeoutError
            except ImportError as err:
                with _oauth_flow_lock:
                    _oauth_flow_state["status"] = "error"
                    _oauth_flow_state["error"] = "google_auth_oauthlib module not installed in Python environment."
                print(f"[OAuth2 ERROR] google_auth_oauthlib missing: {err}", file=sys.stderr, flush=True)
                return

            import requests

            if not OAUTH_CLIENT_SECRETS_FILE.exists():
                with _oauth_flow_lock:
                    _oauth_flow_state["status"] = "error"
                    _oauth_flow_state["error"] = "oauth_client_secrets.json file missing. Upload it first."
                print("[OAuth2 ERROR] oauth_client_secrets.json not found.", file=sys.stderr, flush=True)
                return

            print(f"[OAuth2] Reading credentials from {OAUTH_CLIENT_SECRETS_FILE.name}...", flush=True)
            flow = InstalledAppFlow.from_client_secrets_file(
                str(OAUTH_CLIENT_SECRETS_FILE),
                scopes=DRIVE_SCOPES,
            )

            print("[OAuth2] Launching local HTTP receiver (port=0, timeout=300s) and opening browser...", flush=True)
            try:
                creds = flow.run_local_server(
                    host="localhost",
                    port=0,
                    open_browser=True,
                    timeout_seconds=300,
                    authorization_prompt_message="[OAuth2] Please visit this URL to authorize Google Drive: {url}",
                    success_message="Google Drive authorization complete! You may close this browser window and return to the dashboard."
                )
            except WSGITimeoutError:
                with _oauth_flow_lock:
                    _oauth_flow_state["status"] = "error"
                    _oauth_flow_state["error"] = "Authorization timed out (5 minute limit). Please click Authorize with Google again."
                print("[OAuth2 TIMEOUT] User did not complete login within 5 minutes.", file=sys.stderr, flush=True)
                return

            print("[OAuth2] Authorization code received! Fetching user account info...", flush=True)
            user_email = None
            try:
                resp = requests.get(
                    "https://www.googleapis.com/oauth2/v3/userinfo",
                    headers={"Authorization": f"Bearer {creds.token}"},
                    timeout=10
                )
                if resp.ok:
                    user_email = resp.json().get("email")
            except Exception as e:
                print(f"[OAuth2 WARN] Failed to fetch user profile info: {e}", flush=True)

            token_data = {
                "token": creds.token,
                "refresh_token": creds.refresh_token,
                "token_uri": creds.token_uri,
                "client_id": creds.client_id,
                "client_secret": creds.client_secret,
                "scopes": list(creds.scopes),
                "user_email": user_email,
            }
            OAUTH_TOKEN_FILE.write_text(json.dumps(token_data, indent=2), encoding="utf-8")
            
            with _oauth_flow_lock:
                _oauth_flow_state["status"] = "done"
                _oauth_flow_state["error"] = None
            
            logger.info(f"OAuth2 Drive authorization completed successfully for user: {user_email}")
            print(f"[OAuth2 SUCCESS] Token saved to {OAUTH_TOKEN_FILE.name}. Authorized user: {user_email}", flush=True)

        except Exception as exc:
            import traceback
            tb_str = traceback.format_exc()
            with _oauth_flow_lock:
                _oauth_flow_state["status"] = "error"
                _oauth_flow_state["error"] = f"{type(exc).__name__}: {exc}"
            logger.error(f"OAuth2 Drive authorization failed: {exc}\n{tb_str}")
            print(f"[OAuth2 ERROR] Exception occurred:\n{tb_str}", file=sys.stderr, flush=True)

    t = threading.Thread(target=_run, daemon=True)
    t.start()


def get_oauth_flow_status() -> dict:
    with _oauth_flow_lock:
        return dict(_oauth_flow_state)



def revoke_oauth_token() -> None:
    """Delete the saved OAuth2 token (sign out from Google Drive OAuth)."""
    try:
        if OAUTH_TOKEN_FILE.exists():
            # Optionally revoke the token via Google
            data = json.loads(OAUTH_TOKEN_FILE.read_text(encoding="utf-8"))
            token = data.get("refresh_token") or data.get("token")
            if token:
                try:
                    import requests
                    requests.post("https://oauth2.googleapis.com/revoke", params={"token": token}, timeout=5)
                except Exception:
                    pass
            OAUTH_TOKEN_FILE.unlink()
        _oauth_flow_state["status"] = "idle"
        _oauth_flow_state["error"] = None
    except Exception as exc:
        logger.warning(f"Failed to revoke OAuth token: {exc}")


# ─── Service Account helpers (kept for backward-compat / testing) ─────────────

def find_credentials_file() -> Optional[Path]:
    """Search for service_account.json in common locations."""
    candidates = [
        ROOT / "service_account.json",
        ROOT / "credentials.json",
        ROOT / "backend" / "service_account.json",
        ROOT / "backend" / "credentials.json",
        ROOT / "sample_data" / "service_account.json",
        ROOT / "sample_data" / "credentials.json",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def get_service_account_email() -> Optional[str]:
    """Return the client_email from service_account.json if present."""
    cred_path = find_credentials_file()
    if not cred_path:
        return None
    try:
        data = json.loads(cred_path.read_text(encoding="utf-8"))
        return data.get("client_email")
    except Exception:
        return None


def extract_folder_id(folder_input: Optional[str]) -> Optional[str]:
    """Extract folder ID from a raw ID or full Google Drive folder URL."""
    if not folder_input or not str(folder_input).strip():
        return None
    s = str(folder_input).strip()
    if "drive.google.com" in s and "/folders/" in s:
        parts = s.split("/folders/")
        if len(parts) > 1:
            return parts[1].split("?")[0].split("/")[0].strip()
    return s


def is_drive_available() -> bool:
    """Return True if ANY Drive method is available (OAuth2 preferred, service account fallback)."""
    try:
        import googleapiclient.discovery
    except ImportError:
        return False
    return is_oauth_drive_available() or find_credentials_file() is not None


def get_drive_service():
    """Return an authenticated Drive service — OAuth2 user credentials preferred."""
    if is_oauth_drive_available():
        return get_oauth_drive_service()
    # Fall back to service account
    cred_path = find_credentials_file()
    if not cred_path:
        raise FileNotFoundError("No Google Drive credentials found. Set up OAuth2 or upload service_account.json.")
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(str(cred_path), scopes=DRIVE_SCOPES)
    return build("drive", "v3", credentials=creds)


# ─── Drive operations ──────────────────────────────────────────────────────────

def create_drive_folder(service, folder_name: str, parent_id: Optional[str] = None) -> str:
    """Create a folder on Google Drive and return its folder ID."""
    file_metadata: dict = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        file_metadata["parents"] = [parent_id]

    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    if parent_id:
        query += f" and '{parent_id}' in parents"

    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get("files", [])
    if items:
        return items[0]["id"]

    folder = service.files().create(body=file_metadata, fields="id").execute()
    return folder.get("id")


def set_public_read_permission(service, file_or_folder_id: str):
    """Set 'anyone with link can view' permission on a Google Drive file or folder."""
    try:
        service.permissions().create(
            fileId=file_or_folder_id,
            body={"type": "anyone", "role": "reader"},
            fields="id",
        ).execute()
    except Exception as exc:
        logger.warning(f"Could not set public permissions for {file_or_folder_id}: {exc}")


def upload_file_to_drive(service, file_path: str, parent_folder_id: str, mime_type: str = "application/pdf") -> str:
    """Upload a file to Google Drive under parent_folder_id and return file ID."""
    from googleapiclient.http import MediaFileUpload

    file_name = os.path.basename(file_path)

    query = f"name = '{file_name}' and '{parent_folder_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get("files", [])
    if items:
        return items[0]["id"]

    file_metadata = {"name": file_name, "parents": [parent_folder_id]}
    media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
    uploaded_file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
    ).execute()
    return uploaded_file.get("id")


def upload_poc_certificates_to_drive(
    poc_name: str,
    poc_email: str,
    group_df,
    pdf_paths: List[str],
    root_folder_name: str = "ANVESHA 2026 Certificates",
    parent_folder_id: Optional[str] = None,
    active_type: str = "participation",
) -> Optional[Dict[str, Any]]:
    """
    Upload certificates to Google Drive using OAuth2 (preferred) or service account.
    Creates POC folder hierarchy, organizes by event (and position if multiple winners exist), sets view permissions.
    Returns URLs for the POC folder and individual event subfolders.
    """
    if not is_drive_available():
        logger.info("Google Drive not available — no credentials configured.")
        return None

    using_oauth = is_oauth_drive_available()
    try:
        service = get_drive_service()

        # Resolve root folder
        clean_parent_id = extract_folder_id(parent_folder_id)

        if clean_parent_id:
            root_id = clean_parent_id
        else:
            root_id = create_drive_folder(service, root_folder_name)

        # Create POC folder (only this folder is shared publicly)
        safe_poc_str = poc_name if poc_name else str(poc_email).split("@")[0]
        poc_folder_name = f"Certificates - {safe_poc_str} ({poc_email})"
        poc_folder_id = create_drive_folder(service, poc_folder_name, parent_id=root_id)
        set_public_read_permission(service, poc_folder_id)

        poc_folder_url = f"https://drive.google.com/drive/folders/{poc_folder_id}"
        event_folder_urls: Dict[str, str] = {}

        is_results = str(active_type).lower() in ["results", "winner", "merit"]

        # Upload PDFs organized by Event Name (and Position if multiple winners exist in results mode)
        if group_df is not None and "event_name" in group_df.columns:
            import pandas as pd

            def format_pos_folder(p):
                p_str = str(p).strip()
                if not p_str or pd.isna(p) or p_str.lower() in ["nan", "none"]:
                    return ""
                p_lower = p_str.lower()
                if "1" in p_lower or "first" in p_lower:
                    return "1st Place"
                if "2" in p_lower or "second" in p_lower:
                    return "2nd Place"
                if "3" in p_lower or "third" in p_lower:
                    return "3rd Place"
                return p_str

            for raw_ev, ev_group in group_df.groupby("event_name", dropna=False):
                ev_name = str(raw_ev).strip() if raw_ev and not pd.isna(raw_ev) and str(raw_ev).strip() else "General Event"
                ev_folder_id = create_drive_folder(service, ev_name, parent_id=poc_folder_id)
                set_public_read_permission(service, ev_folder_id)
                event_folder_urls[ev_name] = f"https://drive.google.com/drive/folders/{ev_folder_id}"

                # Check if multiple winning positions exist for this event in results mode
                has_positions = is_results and "position" in ev_group.columns
                if has_positions:
                    formatted_positions = ev_group["position"].apply(format_pos_folder)
                    unique_positions = [p for p in formatted_positions.unique() if p]
                else:
                    unique_positions = []

                if has_positions and len(unique_positions) > 1:
                    # Multiple winning teams/positions -> Create Position subfolders inside Event folder
                    pos_folders: Dict[str, str] = {}
                    for pos_name in unique_positions:
                        pos_id = create_drive_folder(service, pos_name, parent_id=ev_folder_id)
                        set_public_read_permission(service, pos_id)
                        pos_folders[pos_name] = pos_id

                    for idx, row in ev_group.iterrows():
                        pdf_p = row.get("pdf_path")
                        pos_val = format_pos_folder(row.get("position"))
                        target_folder_id = pos_folders.get(pos_val, ev_folder_id)
                        if pdf_p and os.path.exists(pdf_p):
                            upload_file_to_drive(service, pdf_p, target_folder_id, mime_type="application/pdf")
                else:
                    # Single winning position or participation mode -> Dump directly into Event folder
                    for _, row in ev_group.iterrows():
                        pdf_p = row.get("pdf_path")
                        if pdf_p and os.path.exists(pdf_p):
                            upload_file_to_drive(service, pdf_p, ev_folder_id, mime_type="application/pdf")
        else:
            for pdf_p in pdf_paths:
                if pdf_p and os.path.exists(pdf_p):
                    upload_file_to_drive(service, pdf_p, poc_folder_id, mime_type="application/pdf")

        mode = "OAuth2 (your Google account)" if using_oauth else "Service Account"
        logger.info(f"Drive upload complete via {mode} for {poc_email}")
        return {
            "poc_folder_url": poc_folder_url,
            "event_folder_urls": event_folder_urls,
        }

    except Exception as exc:
        logger.error(f"Error uploading to Google Drive for {poc_email}: {exc}")
        raise exc
