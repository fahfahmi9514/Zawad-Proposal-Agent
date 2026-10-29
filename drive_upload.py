"""
drive_upload.py — Upload generated proposals to a shared Google Drive folder.
Uses a service account JSON key (no browser login needed on any laptop).

One-time setup (see README or ask your admin):
  1. Google Cloud Console → create project → enable Drive API
  2. Create a service account → download JSON key → save as config/drive_credentials.json
  3. Share your Google Drive folder with the service account email (Editor access)
  4. Copy the folder ID from the Drive URL and add to config/config.json:
       "drive_folder_id": "1AbCdEfGhIjKlMnOpQrStUvWxYz"
       "drive_credentials": "config/drive_credentials.json"
"""

import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def upload_proposal(file_path: str, config: dict) -> str | None:
    """
    Upload a PPTX file to the shared Google Drive folder.
    Returns the file's shareable Drive link, or None if Drive is not configured or upload fails.
    Silent on failure — never crashes the chatbot.
    """
    folder_id  = config.get('drive_folder_id', '').strip()
    creds_path = config.get('drive_credentials', '').strip()

    if not folder_id or not creds_path:
        return None  # Drive not configured — skip silently

    # Resolve relative paths from the project root
    cp = Path(creds_path)
    if not cp.is_absolute():
        cp = Path(__file__).parent / cp
    if not cp.exists():
        logger.warning(f'Drive credentials not found: {cp}')
        return None

    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_file(
            str(cp),
            scopes=['https://www.googleapis.com/auth/drive.file']
        )
        service = build('drive', 'v3', credentials=creds, cache_discovery=False)

        file_name = Path(file_path).name
        metadata  = {'name': file_name, 'parents': [folder_id]}
        media     = MediaFileUpload(
            file_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            resumable=True
        )
        result = service.files().create(
            body=metadata, media_body=media, fields='id,webViewLink'
        ).execute()

        link = result.get('webViewLink', '')
        logger.info(f'Drive upload OK: {file_name} → {link}')
        return link

    except ImportError:
        logger.warning('google-api-python-client not installed — Drive upload skipped')
        return None
    except Exception as e:
        logger.warning(f'Drive upload failed: {e}')
        return None
