"""
Unit Tests for OneDrive / SharePoint Download with Mocked HTTP Responses
========================================================================
Validates all components of test_onedrive_download.py:
  1. Signature validation (OpenXML, HTML login, expired links, corrupt files)
  2. Candidate URL generation across all SharePoint / OneDrive formats
  3. HTML parsing for embedded download endpoints
  4. Direct and fallback download flows using mocked HTTP responses
  5. Error classification (Auth required, Link expired, HTTP errors)
  6. Secret redaction and log sanitization
  7. End-to-end main() execution
"""

import io
import os
import unittest
import urllib.error
import zipfile
from unittest.mock import MagicMock, patch
import pandas as pd

import test_onedrive_download as td


def create_mock_xlsx_bytes(sheet_name: str = "in") -> bytes:
    """Generates an in-memory valid OpenXML (.xlsx) file."""
    buf = io.BytesIO()
    df = pd.DataFrame({
        "NAME": ["Test Client A", "Test Client B"],
        "STATUS": ["DEMO_SCHEDULED", "COMPLETED"],
        "DEMO_DATE": ["2026-11-14", "2026-11-15"],
    })
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def create_non_excel_zip_bytes() -> bytes:
    """Generates an in-memory ZIP file that does not have OpenXML structure."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("test.txt", "This is just a text file inside a zip.")
    return buf.getvalue()


class TestOneDriveDownloadMock(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.valid_xlsx = create_mock_xlsx_bytes("in")

    # -------------------------------------------------------------------------
    # 1. Signature & Content Validation Tests
    # -------------------------------------------------------------------------
    def test_signature_valid_xlsx(self):
        valid, msg = td.validate_excel_signature(self.valid_xlsx)
        self.assertTrue(valid)
        self.assertIn("Valid Excel", msg)

    def test_signature_rejects_empty(self):
        valid, msg = td.validate_excel_signature(b"")
        self.assertFalse(valid)
        self.assertIn("empty", msg.lower())

    def test_signature_detects_login_html(self):
        login_html = (
            b"<!DOCTYPE html><html><head><title>Sign in</title></head>"
            b"<body>Sign in to your account at login.microsoftonline.com</body></html>"
        )
        valid, msg = td.validate_excel_signature(login_html)
        self.assertFalse(valid)
        self.assertIn("Authentication Required", msg)

    def test_signature_detects_expired_html(self):
        expired_html = (
            b"<!DOCTYPE html><html><body>This link has expired or was removed by the owner.</body></html>"
        )
        valid, msg = td.validate_excel_signature(expired_html)
        self.assertFalse(valid)
        self.assertIn("Link Expired", msg)

    def test_signature_detects_generic_html_viewer(self):
        viewer_html = (
            b"<!DOCTYPE html><html><body><div id='OneNoteApp'>Office Online Viewer</div></body></html>"
        )
        valid, msg = td.validate_excel_signature(viewer_html)
        self.assertFalse(valid)
        self.assertIn("HTML Viewer Response", msg)

    def test_signature_rejects_invalid_magic_bytes(self):
        garbage = b"\x00\x01\x02\x03\x04\x05Some random binary data"
        valid, msg = td.validate_excel_signature(garbage)
        self.assertFalse(valid)
        self.assertIn("Invalid File Signature", msg)

    def test_signature_rejects_non_openxml_zip(self):
        non_excel_zip = create_non_excel_zip_bytes()
        valid, msg = td.validate_excel_signature(non_excel_zip)
        self.assertFalse(valid)
        self.assertIn("Corrupt File", msg)

    # -------------------------------------------------------------------------
    # 2. Candidate URL Generation Tests
    # -------------------------------------------------------------------------
    def test_candidates_sharepoint_personal(self):
        url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123?e=abc"
        cands = td.generate_download_candidates(url)
        self.assertTrue(any("/personal/user_org/_layouts/15/download.aspx?share=Eb8Token123" in c for c in cands))
        self.assertTrue(any("/personal/user_org/_layouts/15/guestaccess.aspx?share=Eb8Token123&download=1" in c for c in cands))
        self.assertTrue(any("download=1" in c for c in cands))

    def test_candidates_sharepoint_sites(self):
        url = "https://tenant.sharepoint.com/:x:/g/sites/Marketing/Eb8SiteToken?e=xyz"
        cands = td.generate_download_candidates(url)
        self.assertTrue(any("/sites/Marketing/_layouts/15/download.aspx?share=Eb8SiteToken" in c for c in cands))
        self.assertTrue(any("/sites/Marketing/_layouts/15/guestaccess.aspx?share=Eb8SiteToken&download=1" in c for c in cands))

    def test_candidates_sourcedoc(self):
        url = (
            "https://tenant-my.sharepoint.com/personal/user_org/_layouts/15/Doc.aspx"
            "?sourcedoc={12345678-ABCD-1234-ABCD-1234567890AB}&action=default"
        )
        cands = td.generate_download_candidates(url)
        self.assertTrue(any("UniqueId=12345678-ABCD-1234-ABCD-1234567890AB" in c for c in cands))
        self.assertTrue(any("sourcedoc=12345678-ABCD-1234-ABCD-1234567890AB" in c for c in cands))

    def test_candidates_onedrive_live(self):
        url = "https://onedrive.live.com/view.aspx?cid=123&resid=456"
        cands = td.generate_download_candidates(url)
        self.assertTrue(any("download" in c for c in cands))

    # -------------------------------------------------------------------------
    # 3. HTML Embedded Link Extraction Tests
    # -------------------------------------------------------------------------
    def test_html_link_extraction(self):
        base_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token"
        html = """
        <html>
        <head><title>Viewer</title></head>
        <body>
            <a href="/_layouts/15/download.aspx?share=Eb8Token&amp;source=nav">Download Link</a>
            <script>
            var fileConfig = {
                "downloadUrl": "https://tenant-my.sharepoint.com/personal/user_org/_layouts/15/download.aspx?UniqueId=98765432-1234-5678-90ab-cdef12345678",
                "sourcedoc": "{98765432-1234-5678-90ab-cdef12345678}"
            };
            </script>
        </body>
        </html>
        """
        extracted = td.extract_download_links_from_html(html, base_url)
        self.assertGreater(len(extracted), 0)
        self.assertTrue(any("download.aspx" in u for u in extracted))
        self.assertTrue(any("98765432-1234-5678-90ab-cdef12345678" in u for u in extracted))

    # -------------------------------------------------------------------------
    # 4. Mocked HTTP Download Execution Tests
    # -------------------------------------------------------------------------
    def test_download_direct_success(self):
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"

        mock_resp = MagicMock()
        mock_resp.read.return_value = self.valid_xlsx
        mock_resp.__enter__.return_value = mock_resp

        mock_opener = MagicMock()
        mock_opener.open.return_value = mock_resp

        data = td.download_file_bytes(share_url, opener=mock_opener)
        self.assertEqual(data, self.valid_xlsx)
        self.assertTrue(mock_opener.open.called)

    def test_download_via_html_fallback(self):
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"

        viewer_html = b"""
        <html><body>
        <script>
        var info = {"downloadUrl": "https://tenant-my.sharepoint.com/personal/user_org/_layouts/15/download.aspx?share=Eb8Token123"};
        </script>
        </body></html>
        """

        # First request returns HTML viewer; second request returns valid binary XLSX
        resp_html = MagicMock()
        resp_html.read.return_value = viewer_html
        resp_html.__enter__.return_value = resp_html

        resp_xlsx = MagicMock()
        resp_xlsx.read.return_value = self.valid_xlsx
        resp_xlsx.__enter__.return_value = resp_xlsx

        mock_opener = MagicMock()
        mock_opener.open.side_effect = [resp_html, resp_xlsx]

        data = td.download_file_bytes(share_url, opener=mock_opener)
        self.assertEqual(data, self.valid_xlsx)
        self.assertEqual(mock_opener.open.call_count, 2)

    def test_download_fails_when_auth_required(self):
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"

        login_html = b"<!DOCTYPE html><html><body>Sign in to your account at login.microsoftonline.com</body></html>"
        mock_resp = MagicMock()
        mock_resp.read.return_value = login_html
        mock_resp.__enter__.return_value = mock_resp

        mock_opener = MagicMock()
        mock_opener.open.return_value = mock_resp

        with self.assertRaises(RuntimeError) as ctx:
            td.download_file_bytes(share_url, opener=mock_opener)

        self.assertIn("Authentication Required", str(ctx.exception))
        # Ensure no secret URL leaked
        self.assertNotIn("Eb8Token123", str(ctx.exception))

    def test_download_fails_when_link_expired(self):
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"

        expired_html = b"<!DOCTYPE html><html><body>This link has expired or the file was deleted.</body></html>"
        mock_resp = MagicMock()
        mock_resp.read.return_value = expired_html
        mock_resp.__enter__.return_value = mock_resp

        mock_opener = MagicMock()
        mock_opener.open.return_value = mock_resp

        with self.assertRaises(RuntimeError) as ctx:
            td.download_file_bytes(share_url, opener=mock_opener)

        self.assertIn("Link Expired or Revoked", str(ctx.exception))

    def test_download_http_error_handling(self):
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"

        err_fp = io.BytesIO(b"Access Denied")
        mock_opener = MagicMock()
        mock_opener.open.side_effect = urllib.error.HTTPError(
            url="https://tenant-my.sharepoint.com/test",
            code=403,
            msg="Forbidden",
            hdrs={},
            fp=err_fp,
        )

        with self.assertRaises(RuntimeError) as ctx:
            td.download_file_bytes(share_url, opener=mock_opener)

        err_fp.close()
        self.assertIn("Access Denied", str(ctx.exception))
        self.assertNotIn("Eb8Token123", str(ctx.exception))

    # -------------------------------------------------------------------------
    # 5. Redaction and Sanitization Tests
    # -------------------------------------------------------------------------
    def test_sanitization(self):
        secret = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8SecretToken123?e=secretValue"
        raw_error = f"Error connecting to {secret} with code=XYZ123 and auth=SECRET_KEY, also see https://example.com/status"
        clean = td.sanitize_text(raw_error, secret)

        self.assertNotIn("Eb8SecretToken123", clean)
        self.assertNotIn("secretValue", clean)
        self.assertNotIn("XYZ123", clean)
        self.assertNotIn("SECRET_KEY", clean)
        self.assertIn("[ONEDRIVE_SHARE_URL_REDACTED]", clean)
        self.assertIn("[URL_REDACTED]", clean)

    # -------------------------------------------------------------------------
    # 6. Main Flow Tests
    # -------------------------------------------------------------------------
    def test_main_missing_env(self):
        with patch.dict(os.environ, {}, clear=True):
            with patch("sys.exit", side_effect=SystemExit) as mock_exit:
                with self.assertRaises(SystemExit):
                    td.main()
                mock_exit.assert_called_with(1)

    def test_main_success_flow(self):
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"
        with patch.dict(os.environ, {"ONEDRIVE_SHARE_URL": share_url}):
            with patch.object(td, "download_file_bytes", return_value=self.valid_xlsx):
                with patch("sys.exit", side_effect=SystemExit) as mock_exit:
                    with self.assertRaises(SystemExit):
                        td.main()
                    mock_exit.assert_called_with(0)

    def test_main_missing_sheet_in(self):
        xlsx_wrong_sheet = create_mock_xlsx_bytes("OtherSheet")
        share_url = "https://tenant-my.sharepoint.com/:x:/g/personal/user_org/Eb8Token123"
        with patch.dict(os.environ, {"ONEDRIVE_SHARE_URL": share_url}):
            with patch.object(td, "download_file_bytes", return_value=xlsx_wrong_sheet):
                with patch("sys.exit", side_effect=SystemExit) as mock_exit:
                    with self.assertRaises(SystemExit):
                        td.main()
                    mock_exit.assert_called_with(1)


if __name__ == "__main__":
    unittest.main()
