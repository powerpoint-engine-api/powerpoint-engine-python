"""PowerPoint Engine API client.

Thin wrapper over the public REST API at https://powerpointengine.io.
All methods return the parsed JSON response; generated files are fetched
from the signed ``downloadUrl`` in the result (valid for 24 hours).
"""

import json
import mimetypes
import os
from typing import Any, Dict, List, Optional

import requests

from .exceptions import (
    AuthenticationError,
    NotFoundError,
    PowerPointEngineError,
    RateLimitError,
    ServerError,
    ValidationError,
)

DEFAULT_BASE_URL = "https://powerpointengine.io"

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


def _file_part(path):
    """(filename, handle, mime) tuple so the API sees the right content type."""
    mime = mimetypes.guess_type(path)[0] or PPTX_MIME
    return (os.path.basename(path), open(path, "rb"), mime)


class PowerPointEngine:
    """Client for the PowerPoint Engine API.

    Args:
        session_id: Optional account id (the ``sessionId`` the dashboard
            shows). Ties generations to your account for credits and history;
            anonymous calls work but return watermarked files.
        base_url: API origin (default: https://powerpointengine.io).
        timeout: Per-request timeout in seconds.
    """

    def __init__(
        self,
        session_id: Optional[str] = None,
        base_url: str = DEFAULT_BASE_URL,
        timeout: int = 120,
    ):
        self.session_id = session_id
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers["User-Agent"] = "powerpoint-engine-python/2.0.0"

    # -- generation ---------------------------------------------------------

    def generate(
        self,
        markup: Optional[str] = None,
        template: Optional[Dict[str, Any]] = None,
        theme: str = "corporate",
        brand: Optional[Dict[str, str]] = None,
        font: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Generate a .pptx from markup text or a structured template.

        Exactly one of ``markup`` / ``template`` is required. ``markup`` is
        the Markdown dialect (``# title``, ``## slide``, bullets, tables,
        ```` ```chart ```` blocks). ``brand`` is a dict of 6-hex colors like
        ``{"primary": "#E4002B"}``; ``font`` overrides the theme font.
        """
        if not markup and not template:
            raise ValueError("either markup or template is required")
        body: Dict[str, Any] = {"theme": theme}
        if markup:
            body["markup"] = markup
        if template:
            body["template"] = template
        if brand:
            body["brand"] = brand
        if font:
            body["font"] = font
        if self.session_id:
            body["sessionId"] = self.session_id
        return self._request_json("/api/powerpoint/generate", body)

    # -- operations on an existing .pptx ------------------------------------

    def replace(
        self,
        file_path: str,
        replacements: Dict[str, str],
        replace_mode: str = "placeholders",
        images: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Replace text (and optionally images) in your own .pptx in place.

        ``replace_mode`` is ``placeholders`` ({{key}} tokens) or ``objects``
        (match shapes by name). ``images`` maps a shape name or alt text to a
        local image path; position, size and effects are kept.
        """
        data = {
            "replacements": json.dumps(replacements),
            "replaceMode": replace_mode,
        }
        files = {"file": _file_part(file_path)}
        try:
            if images:
                for shape_name, image_path in images.items():
                    files[f"image:{shape_name}"] = _file_part(image_path)
            return self._request_multipart("/api/powerpoint/replace", data, files)
        finally:
            for part in files.values():
                part[1].close()

    def edit(self, file_path: str, operations: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Duplicate / delete / move slides in your own .pptx.

        ``operations`` apply sequentially, slide numbers are 1-based, e.g.::

            [{"op": "duplicate", "slide": 2},
             {"op": "delete", "slide": 5},
             {"op": "move", "slide": 3, "to": 1}]
        """
        data = {"operations": json.dumps(operations)}
        part = _file_part(file_path)
        try:
            return self._request_multipart(
                "/api/powerpoint/edit", data, {"file": part}
            )
        finally:
            part[1].close()

    def merge(self, file_paths: List[str]) -> Dict[str, Any]:
        """Merge 2-5 decks into one; each keeps its own design."""
        if not 2 <= len(file_paths) <= 5:
            raise ValueError("merge takes 2 to 5 files")
        parts = [_file_part(p) for p in file_paths]
        try:
            files = [("files", part) for part in parts]
            return self._request_multipart("/api/powerpoint/merge", {}, files)
        finally:
            for part in parts:
                part[1].close()

    def to_pdf(self, file_path: str) -> Dict[str, Any]:
        """Convert a .pptx to PDF; result has a signed PDF downloadUrl."""
        part = _file_part(file_path)
        try:
            return self._request_multipart(
                "/api/powerpoint/pdf", {}, {"file": part}
            )
        finally:
            part[1].close()

    def translate(self, file_path: str, target_lang: str) -> Dict[str, Any]:
        """Translate all text in a .pptx in place (layout preserved)."""
        data = {"targetLang": target_lang}
        part = _file_part(file_path)
        try:
            return self._request_multipart(
                "/api/powerpoint/translate", data, {"file": part}
            )
        finally:
            part[1].close()

    # -- helpers -------------------------------------------------------------

    def download(self, result: Dict[str, Any], to_path: str) -> str:
        """Download the generated file from a response's downloadUrl."""
        url = (result.get("result") or {}).get("downloadUrl") or result.get(
            "downloadUrl"
        )
        if not url:
            raise ValueError("response has no downloadUrl")
        response = requests.get(url, timeout=self.timeout)
        response.raise_for_status()
        with open(to_path, "wb") as fh:
            fh.write(response.content)
        return to_path

    def _request_json(self, endpoint: str, body: Dict[str, Any]) -> Dict[str, Any]:
        response = self.session.post(
            self.base_url + endpoint, json=body, timeout=self.timeout
        )
        return self._parse(response)

    def _request_multipart(self, endpoint: str, data, files) -> Dict[str, Any]:
        if self.session_id:
            data = dict(data)
            data["sessionId"] = self.session_id
        response = self.session.post(
            self.base_url + endpoint, data=data, files=files, timeout=self.timeout
        )
        return self._parse(response)

    def _parse(self, response: requests.Response) -> Dict[str, Any]:
        try:
            payload = response.json()
        except ValueError:
            payload = {"error": response.text[:500]}
        if response.ok:
            return payload
        message = payload.get("error", "Unknown error")
        status = response.status_code
        if status == 401:
            raise AuthenticationError(message, status)
        if status == 400:
            raise ValidationError(message, status)
        if status == 404:
            raise NotFoundError(message, status)
        if status == 429:
            raise RateLimitError(message, status)
        if status >= 500:
            raise ServerError(message, status)
        raise PowerPointEngineError(message, status)
