"""
OneDrive Sharing Link Download & Inspection Test
================================================
Validates that 'lumos10.xlsx' can be downloaded from OneDrive / SharePoint
via an anonymous sharing link ('Anyone with the link can view') and parsed
with pandas/openpyxl without exposing secrets or data in logs.

Inputs:
  ONEDRIVE_SHARE_URL environment variable (from GitHub Secrets).

Outputs:
  Reports download status, row count, column count, and column names only.
  Never prints the URL, token, row values, or customer data in logs.
"""

import base64
import http.cookiejar
import io
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import pandas as pd


DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


def sanitize_text(text: str, secret_url: str | None = None) -> str:
    """Masks secret URL, tokens, query parameters, and endpoints from log messages."""
    if not text:
        return ""
    clean = str(text)
    if secret_url and len(secret_url) > 5:
        clean = clean.replace(secret_url, "[ONEDRIVE_SHARE_URL_REDACTED]")
    # Mask query parameters and key-value pairs that could hold tokens or keys
    clean = re.sub(
        r"(\b(?:share|token|code|key|auth|e)=)[^&\s'\"]+",
        r"\1[REDACTED]",
        clean,
        flags=re.IGNORECASE,
    )
    # Mask full http/https URLs to prevent any token/host exposure in exceptions
    clean = re.sub(r"https?://[^\s\"'<>]+", "[URL_REDACTED]", clean)
    return clean


def validate_excel_signature(data: bytes) -> tuple[bool, str]:
    """
    Validates that the downloaded byte sequence is a genuine Excel (.xlsx) file
    and diagnoses HTML error/login pages if returned instead.
    """
    if not data or len(data) < 4:
        return False, "Downloaded content is empty or incomplete (fewer than 4 bytes)."

    if not data.startswith(b"PK"):
        # Check if the response is an HTML page (login, error, or web app)
        head_sample = data[:1000].lower()
        if b"<html" in head_sample or b"<!doctype html" in head_sample:
            text_preview = data[:4000].decode("utf-8", errors="ignore").lower()
            if (
                "sign in to your account" in text_preview
                or "login.microsoftonline.com" in text_preview
                or "login.live.com" in text_preview
                or "aad" in text_preview
            ):
                return False, (
                    "Authentication Required: The OneDrive sharing link requires organizational sign-in. "
                    "Ensure the sharing link is created as 'Anyone with the link can view'."
                )
            if (
                "link has expired" in text_preview
                or "link was removed" in text_preview
                or "item does not exist" in text_preview
            ):
                return False, (
                    "Link Expired or Revoked: The OneDrive sharing link has expired or the file was deleted."
                )
            return False, (
                "HTML Viewer Response: Received SharePoint web viewer HTML instead of binary spreadsheet. "
                "Direct download could not be initiated."
            )
        return False, f"Invalid File Signature: Expected OpenXML (.xlsx) PK header, got: {data[:4]!r}"

    # Verify that the ZIP container contains valid OpenXML structure
    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
            namelist = zf.namelist()
            if not any(n.startswith("xl/") or n == "[Content_Types].xml" for n in namelist):
                return False, "Corrupt File: Archive is a valid ZIP but does not contain Excel OpenXML structures."
    except Exception as e:
        return False, f"Corrupt OpenXML Archive: {e}"

    return True, "Valid Excel .xlsx binary"


def generate_download_candidates(sharing_url: str) -> list[str]:
    """
    Generates prioritized direct-download candidate URLs from a OneDrive / SharePoint link:
      1. SharePoint Online direct download endpoint (_layouts/15/download.aspx?share=...)
      2. SharePoint guest access download endpoint (_layouts/15/guestaccess.aspx?share=...&download=1)
      3. Sourcedoc GUID direct download endpoint (if Doc.aspx with sourcedoc is supplied)
      4. Appending download=1 query parameter
      5. Microsoft Graph Shares API (for public anonymous links)
      6. Original URL as fallback
    """
    candidates = []
    clean_url = sharing_url.strip()
    parsed = urllib.parse.urlparse(clean_url)
    scheme_netloc = f"{parsed.scheme}://{parsed.netloc}"

    # 1. SharePoint guest shortlinks (/:x:/g/..., /:x:/s/..., etc.)
    sp_match = re.match(
        r"^/:[a-z]:/[a-z]/(?:(personal/[^/]+|sites/[^/]+|teams/[^/]+)/)?([^/?#]+)",
        parsed.path,
    )
    if sp_match:
        site_prefix = sp_match.group(1) or ""
        token = sp_match.group(2)
        base_path = f"/{site_prefix}" if site_prefix else ""
        candidates.append(f"{scheme_netloc}{base_path}/_layouts/15/download.aspx?share={token}")
        candidates.append(f"{scheme_netloc}{base_path}/_layouts/15/guestaccess.aspx?share={token}&download=1")

    # 2. Sourcedoc GUID if present in query parameters
    qs = urllib.parse.parse_qs(parsed.query)
    if "sourcedoc" in qs:
        guid = qs["sourcedoc"][0].strip("{}")
        prefix = parsed.path.split("/_layouts/")[0] if "/_layouts/" in parsed.path else ""
        candidates.append(f"{scheme_netloc}{prefix}/_layouts/15/download.aspx?UniqueId={guid}")
        candidates.append(f"{scheme_netloc}{prefix}/_layouts/15/download.aspx?sourcedoc={guid}")

    # 3. Query param download=1
    if "download=1" in clean_url:
        candidates.append(clean_url)
    else:
        sep = "&" if parsed.query else "?"
        dl_url = re.sub(r"[?&]action=[^&]+", "", clean_url)
        candidates.append(f"{dl_url}{sep}download=1")

    # 4. Consumer OneDrive (1drv.ms / onedrive.live.com)
    if "1drv.ms" in parsed.netloc and not clean_url.endswith("download=1"):
        sep = "&" if parsed.query else "?"
        candidates.append(f"{clean_url}{sep}download=1")
    elif "onedrive.live.com" in parsed.netloc and "view.aspx" in parsed.path:
        candidates.append(clean_url.replace("view.aspx", "download"))

    # 5. Microsoft Graph Shares API (public anonymous driveItem/content)
    try:
        raw_b64 = base64.b64encode(clean_url.encode("utf-8")).decode("utf-8")
        b64_token = "u!" + raw_b64.replace("+", "-").replace("/", "_").rstrip("=")
        candidates.append(f"https://graph.microsoft.com/v1.0/shares/{b64_token}/driveItem/content")
    except Exception:
        pass

    # 6. Raw URL
    if clean_url not in candidates:
        candidates.append(clean_url)

    # De-duplicate while preserving order
    seen = set()
    unique_candidates = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique_candidates.append(c)

    return unique_candidates


def extract_download_links_from_html(html_text: str, base_url: str) -> list[str]:
    """
    Parses a SharePoint web viewer HTML response to discover embedded download URLs,
    JSON download attributes, or UniqueId/sourcedoc GUIDs.
    """
    links = []
    parsed = urllib.parse.urlparse(base_url)
    scheme_netloc = f"{parsed.scheme}://{parsed.netloc}"

    # 1. Direct download.aspx links
    for m in re.findall(r'https?://[^"\'\s<>]*/_layouts/15/download\.aspx\?[^"\'\s<>]+', html_text):
        links.append(m.replace("&amp;", "&"))
    for m in re.findall(r'(/_layouts/15/download\.aspx\?[^"\'\s<>]+)', html_text):
        links.append(f"{scheme_netloc}{m.replace('&amp;', '&')}")

    # 2. Extract downloadUrl JSON attribute
    for durl in re.findall(
        r'["\'](?:downloadUrl|DownloadUrl)["\']\s*:\s*["\'](https?://[^"\']+)["\']',
        html_text,
    ):
        links.append(durl.replace("\\/", "/").replace("&amp;", "&"))

    # 3. Extract UniqueId or sourcedoc GUID
    for guid in re.findall(
        r'["\']?(?:UniqueId|sourcedoc)["\']?\s*[:=]\s*["\']?\{?([0-9a-fA-F\-]{36})\}?["\']?',
        html_text,
    ):
        links.append(f"{scheme_netloc}/_layouts/15/download.aspx?UniqueId={guid}")

    # De-duplicate while preserving order
    seen = set()
    unique_links = []
    for link in links:
        if link not in seen:
            seen.add(link)
            unique_links.append(link)

    return unique_links


def create_http_opener() -> urllib.request.OpenerDirector:
    """Creates a cookie-aware HTTP opener that preserves SharePoint guest authentication cookies."""
    cookie_jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cookie_jar),
        urllib.request.HTTPSHandler(),
    )
    return opener


def download_file_bytes(sharing_url: str, opener: urllib.request.OpenerDirector | None = None) -> bytes:
    """
    Downloads raw binary bytes from candidate URLs following redirects and managing cookies.
    If an HTML viewer page is returned, attempts to extract the embedded download endpoint.
    Ensures response matches Excel OpenXML (.xlsx) signature.
    """
    if opener is None:
        opener = create_http_opener()

    candidates = generate_download_candidates(sharing_url)
    diagnostics = []

    for target_url in candidates:
        try:
            req = urllib.request.Request(target_url, headers=DEFAULT_HEADERS, method="GET")
            with opener.open(req, timeout=30) as resp:
                data = resp.read()

                # Validate binary signature
                is_valid, diag_msg = validate_excel_signature(data)
                if is_valid:
                    return data

                diagnostics.append(diag_msg)

                # If HTML was returned, attempt fallback discovery of download link inside HTML
                if "HTML Viewer" in diag_msg or "login" not in diag_msg.lower():
                    html_text = data.decode("utf-8", errors="ignore")
                    extracted_links = extract_download_links_from_html(html_text, target_url)

                    for ext_url in extracted_links:
                        try:
                            ext_req = urllib.request.Request(ext_url, headers=DEFAULT_HEADERS, method="GET")
                            with opener.open(ext_req, timeout=30) as ext_resp:
                                ext_data = ext_resp.read()
                                ext_valid, ext_diag = validate_excel_signature(ext_data)
                                if ext_valid:
                                    return ext_data
                                diagnostics.append(ext_diag)
                        except Exception as e:
                            diagnostics.append(f"Extracted URL fetch error: {type(e).__name__}")

        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                diagnostics.append(
                    f"HTTP {e.code} (Access Denied): Link may require organizational sign-in or specific user permissions."
                )
            elif e.code == 404:
                diagnostics.append(f"HTTP 404 (Not Found): Resource not found at endpoint.")
            else:
                diagnostics.append(f"HTTP {e.code}: {e.reason}")
        except urllib.error.URLError as e:
            diagnostics.append(f"Network error: {e.reason}")
        except Exception as e:
            diagnostics.append(f"Request error: {type(e).__name__}")

    # Prioritize specific actionable diagnostics if present
    auth_errors = [d for d in diagnostics if "Authentication Required" in d or "Access Denied" in d]
    if auth_errors:
        final_cause = auth_errors[0]
    else:
        expired_errors = [d for d in diagnostics if "Expired" in d]
        if expired_errors:
            final_cause = expired_errors[0]
        else:
            final_cause = diagnostics[-1] if diagnostics else "No valid Excel content returned."

    safe_error = sanitize_text(final_cause, sharing_url)
    raise RuntimeError(f"Unable to download Excel workbook via sharing URL. Reason: {safe_error}")


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
        try:
            df = pd.read_excel(stream, sheet_name="in", engine="openpyxl")
        except ValueError as ve:
            if "Worksheet" in str(ve):
                # Try inspecting available sheets safely
                xl_file = pd.ExcelFile(io.BytesIO(raw_bytes), engine="openpyxl")
                print("STATUS: FAILED")
                print(f"ERROR: Worksheet 'in' not found. Available sheets: {xl_file.sheet_names}")
                sys.exit(1)
            raise

        row_count = len(df)
        cols = list(df.columns)

        print("WORKSHEET: 'in'")
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
