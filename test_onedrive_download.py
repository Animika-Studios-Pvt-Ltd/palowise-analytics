"""
OneDrive Sharing Link Download & Inspection Test
================================================
Validates that 'lumos10.xlsx' can be downloaded from OneDrive via a sharing link
and parsed with pandas/openpyxl without exposing secrets or data in logs.

Inputs:
  ONEDRIVE_SHARE_URL environment variable (from GitHub Secrets).

Outputs:
  Reports download status, row count, column count, and column names only.
  Never prints the URL, row values, or customer data in logs.
"""

import base64
import io
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
import pandas as pd


def sanitize_text(text: str, secret_url: str | None) -> str:
    """Masks secret URL and sensitive tokens from error messages."""
    if not text:
        return ""
    clean = str(text)
    if secret_url and len(secret_url) > 5:
        clean = clean.replace(secret_url, "[ONEDRIVE_SHARE_URL_REDACTED]")
    # Mask any lingering token/code parameters
    clean = re.sub(r"[?&](token|code|key|auth|e)=[^&\s]+", r"&\1=[REDACTED]", clean)
    return clean


def get_direct_download_urls(sharing_url: str) -> list[str]:
    """
    Generates candidate download URLs from a OneDrive / SharePoint sharing link:
      1. Standard direct download query (?download=1).
      2. Microsoft Graph Shares API (publicly accessible for 'Anyone' links).
      3. Original URL (if redirect handler can resolve it).
    """
    candidates = []
    clean_url = sharing_url.strip()

    # Candidate 1: Query param download=1 (standard for SharePoint/OneDrive web links)
    if "download=1" in clean_url:
        candidates.append(clean_url)
    else:
        sep = "&" if "?" in clean_url else "?"
        # Replace viewing params if present
        dl_url = re.sub(r"[?&]action=[^&]+", "", clean_url)
        dl_url = f"{dl_url}{sep}download=1"
        candidates.append(dl_url)

    # Candidate 2: Microsoft Graph Shares API
    try:
        # Encode URL to base64url without padding
        raw_b64 = base64.b64encode(clean_url.encode("utf-8")).decode("utf-8")
        b64_token = "u!" + raw_b64.replace("+", "-").replace("/", "_").rstrip("=")
        graph_dl = f"https://graph.microsoft.com/v1.0/shares/{b64_token}/driveItem/content"
        onedrive_api_dl = f"https://api.onedrive.com/v1.0/shares/{b64_token}/root/content"
        candidates.append(graph_dl)
        candidates.append(onedrive_api_dl)
    except Exception:
        pass

    # Candidate 3: Raw link as fallback
    if clean_url not in candidates:
        candidates.append(clean_url)

    return candidates


def download_file_bytes(sharing_url: str) -> bytes:
    """
    Downloads raw binary bytes from candidate URLs following redirects.
    Ensures response is an Excel binary (.xlsx) and not an HTML login page.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
    }

    candidates = get_direct_download_urls(sharing_url)
    last_error = None

    for idx, target_url in enumerate(candidates, 1):
        try:
            req = urllib.request.Request(target_url, headers=headers, method="GET")
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()

                # Verify magic bytes for ZIP/XLSX: starts with b'PK\x03\x04' or b'PK'
                if data.startswith(b"PK"):
                    return data
                elif b"<html" in data[:500].lower() or b"<!doctype html" in data[:500].lower():
                    last_error = (
                        "Downloaded file is an HTML page rather than an Excel binary. "
                        "The OneDrive sharing link may require organizational sign-in or may not be set to "
                        "'Anyone with the link can view'."
                    )
                else:
                    last_error = f"Downloaded file ({len(data):,} bytes) does not match Excel (.xlsx) signature."
        except urllib.error.HTTPError as e:
            last_error = f"HTTP error {e.code}: {e.reason}"
        except urllib.error.URLError as e:
            last_error = f"URL error: {e.reason}"
        except Exception as e:
            last_error = str(e)

    safe_error = sanitize_text(last_error, sharing_url)
    raise RuntimeError(f"Unable to download Excel workbook via sharing URL. Last error: {safe_error}")


def main():
    share_url = os.environ.get("ONEDRIVE_SHARE_URL")

    if not share_url or not share_url.strip():
        print("STATUS: FAILED")
        print("ERROR: Environment variable 'ONEDRIVE_SHARE_URL' is missing or empty.")
        print("Make sure ONEDRIVE_SHARE_URL is configured in GitHub Secrets.")
        sys.exit(1)

    print("Attempting to download workbook from OneDrive sharing link...")

    try:
        raw_bytes = download_file_bytes(share_url)
        print(f"DOWNLOAD STATUS: SUCCESS ({len(raw_bytes):,} bytes received)")

        # Read worksheet 'in'
        stream = io.BytesIO(raw_bytes)
        df = pd.read_excel(stream, sheet_name="in", engine="openpyxl")

        row_count = len(df)
        cols = list(df.columns)

        print(f"WORKSHEET: 'in'")
        print(f"ROW COUNT: {row_count}")
        print(f"COLUMN COUNT: {len(cols)}")
        print(f"COLUMNS: {cols}")
        print("VERIFICATION: PASSED")
        sys.exit(0)

    except Exception as e:
        safe_msg = sanitize_text(str(e), share_url)
        print("STATUS: FAILED")
        print(f"ERROR: {safe_msg}")
        sys.exit(1)


if __name__ == "__main__":
    main()

