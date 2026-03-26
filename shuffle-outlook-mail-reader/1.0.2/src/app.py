#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Mail Reader — Shuffle SOAR App (Microsoft Graph, App-Only Auth)

Provides generic, configurable actions for reading and filtering messages
from any Outlook/Exchange Online mailbox using Microsoft Graph API with
application (client credentials) permissions.

Actions:
    list_messages       — List messages with optional filters
    get_message         — Fetch a single message by its Graph message ID
    extract_with_regex  — Apply a regex pattern to a message body (optional extraction)
"""

import json
import logging
import re
import unicodedata
import datetime as _dt
from typing import Any, Dict, List, Optional, Set

import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Well-known folder names supported by Microsoft Graph
WELL_KNOWN_FOLDERS: Set[str] = {
    "inbox",
    "sentitems",
    "deleteditems",
    "drafts",
    "junkemail",
    "outbox",
    "archive",
    "recoverableitemsdeletions",
}

# ─── Text Normalisation Utilities ────────────────────────────────────────────

# Translation map for Turkish-specific characters to their ASCII equivalents.
# Useful when normalising display names that may contain Turkish letters.
_TR_CHAR_MAP = str.maketrans(
    {
        "ç": "c", "ğ": "g", "ı": "i", "ö": "o", "ş": "s", "ü": "u",
        "Ç": "c", "Ğ": "g", "İ": "i", "I": "i", "Ö": "o", "Ş": "s", "Ü": "u",
    }
)


def _normalise_display_name(name: str) -> str:
    """
    Normalise a display name or username to a lowercase ASCII dot-separated
    identifier (e.g. "Jane Doe" → "jane.doe").

    This is useful when comparing names from email bodies against a configured
    exclusion list, regardless of diacritic or casing differences.
    """
    name = (name or "").strip()
    if not name:
        return ""

    # Treat names with spaces as "First Last" → use first and last token only.
    if " " in name:
        tokens = re.findall(r"[A-Za-z\u00c0-\u024f\u0130\u0131]+", name)
        use = f"{tokens[0]} {tokens[-1]}" if len(tokens) >= 2 else name
    else:
        # Already in username format (e.g. "jane.doe")
        use = name

    use = use.translate(_TR_CHAR_MAP).lower()
    try:
        use = unicodedata.normalize("NFKD", use).encode("ascii", "ignore").decode("ascii")
    except Exception:
        pass

    use = re.sub(r"[^a-z0-9 .]+", "", use)
    use = re.sub(r"\s+", " ", use).strip()
    use = use.replace(" ", ".")
    use = re.sub(r"\.+", ".", use).strip(".")
    return use


def _parse_name_list(raw: str) -> Set[str]:
    """
    Parse a comma/semicolon/newline-separated string of names or usernames
    into a normalised set of identifiers.

    Example:
        "Jane Doe; john.smith" → {"jane.doe", "john.smith"}
    """
    if not raw:
        return set()
    parts = re.split(r"[,\n;]+", raw)
    return {key for p in parts if (key := _normalise_display_name(p))}


# ─── App Class ────────────────────────────────────────────────────────────────

class OutlookMailReader(AppBase):
    """
    Shuffle SOAR app for reading Outlook/Exchange Online mailboxes via
    Microsoft Graph API using application (app-only) credentials.

    All actions require an Azure AD application registration with the
    ``Mail.Read`` application permission granted and admin-consented.
    """

    __version__ = "2.0.0"
    app_name = "Outlook Mail Reader"

    def __init__(self, redis=None, logger=None, **kwargs):
        super().__init__(redis=redis, logger=logger, **kwargs)

    # ─── Private: Authentication ──────────────────────────────────────────

    def _get_access_token(
        self, tenant_id: str, client_id: str, client_secret: str
    ) -> str:
        """
        Obtain an OAuth2 client-credentials access token from Azure AD.

        Args:
            tenant_id:     Azure AD tenant ID (GUID or domain name).
            client_id:     Application (client) ID of the app registration.
            client_secret: Client secret for the app registration.

        Returns:
            A Bearer access token string.

        Raises:
            requests.HTTPError: If the token request fails.
        """
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        response = requests.post(url, data=payload, timeout=30)
        response.raise_for_status()
        return response.json()["access_token"]

    # ─── Private: HTTP helpers ────────────────────────────────────────────

    def _graph_get(
        self,
        url: str,
        token: str,
        params: Optional[Dict[str, Any]] = None,
        prefer_text_body: bool = True,
    ) -> Dict[str, Any]:
        """
        Perform an authenticated GET request against the Microsoft Graph API.

        Args:
            url:              Full Graph API URL.
            token:            Bearer access token.
            params:           Optional OData query parameters.
            prefer_text_body: When True, requests plain-text body content
                              instead of HTML via the ``Prefer`` header.

        Returns:
            Parsed JSON response as a dictionary.

        Raises:
            requests.HTTPError: If the request returns a non-2xx status.
        """
        headers: Dict[str, str] = {"Authorization": f"Bearer {token}"}
        if prefer_text_body:
            headers["Prefer"] = 'outlook.body-content-type="text"'

        log = self.logger or logging.getLogger(__name__)
        try:
            prepped = requests.Request("GET", url, params=params).prepare()
            log.info("[Graph GET] url=%s", prepped.url)
        except Exception:
            pass

        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        return response.json()

    # ─── Private: Folder resolution ───────────────────────────────────────

    def _resolve_folder_url(
        self, mailbox: str, folder: str, token: str
    ) -> str:
        """
        Resolve a folder name or well-known folder name to a Graph messages URL.

        Well-known names (``inbox``, ``sentitems``, ``deleteditems``, etc.) are
        used directly. Other names are looked up via the ``mailFolders`` endpoint
        and matched by ``displayName`` (case-insensitive).

        Args:
            mailbox: The target mailbox user principal name or object ID.
            folder:  Folder display name or well-known name. Defaults to "inbox".
            token:   Bearer access token.

        Returns:
            A Graph API URL ending in ``/messages``.

        Raises:
            ValueError: If a named folder cannot be found in the mailbox.
        """
        folder_clean = (folder or "inbox").strip().lower()

        if folder_clean in WELL_KNOWN_FOLDERS:
            return f"{GRAPH_BASE}/users/{mailbox}/mailFolders/{folder_clean}/messages"

        # Look up by display name
        lookup_url = f"{GRAPH_BASE}/users/{mailbox}/mailFolders"
        data = self._graph_get(
            lookup_url,
            token,
            params={"$select": "id,displayName", "$top": 100},
            prefer_text_body=False,
        )
        folders = data.get("value", [])
        for f in folders:
            if (f.get("displayName") or "").lower() == folder_clean:
                folder_id = f["id"]
                return f"{GRAPH_BASE}/users/{mailbox}/mailFolders/{folder_id}/messages"

        raise ValueError(
            f"Mail folder '{folder}' not found in mailbox '{mailbox}'. "
            f"Available folders: {[f.get('displayName') for f in folders]}"
        )

    # ─── Private: Filter builder ──────────────────────────────────────────

    @staticmethod
    def _build_odata_filter(
        subject_filter: Optional[str] = None,
        sender_filter: Optional[str] = None,
        unread_only: bool = False,
        received_after: Optional[str] = None,
        received_before: Optional[str] = None,
    ) -> Optional[str]:
        """
        Build an OData ``$filter`` expression from optional filter arguments.

        All provided filters are combined with ``and``. An anchor clause
        (``receivedDateTime ge 1900-01-01T00:00:00Z``) is always included
        so that Graph returns results in a consistent order when ``$orderby``
        is also applied.

        Args:
            subject_filter:  Exact subject string match.
            sender_filter:   Sender email address match.
            unread_only:     If True, only return unread messages.
            received_after:  ISO 8601 datetime string (inclusive lower bound).
            received_before: ISO 8601 datetime string (exclusive upper bound).

        Returns:
            An OData filter string, or None if no filters were provided.
        """
        clauses: List[str] = ["receivedDateTime ge 1900-01-01T00:00:00Z"]

        if subject_filter:
            safe = subject_filter.replace("'", "''")
            clauses.append(f"subject eq '{safe}'")

        if sender_filter:
            safe = sender_filter.replace("'", "''")
            clauses.append(f"sender/emailAddress/address eq '{safe}'")

        if unread_only:
            clauses.append("isRead eq false")

        if received_after:
            clauses.append(f"receivedDateTime ge {received_after}")

        if received_before:
            clauses.append(f"receivedDateTime lt {received_before}")

        return " and ".join(clauses)

    # ─── Private: Body helpers ────────────────────────────────────────────

    @staticmethod
    def _extract_body_text(message: Dict[str, Any]) -> str:
        """
        Extract the plain-text body from a Graph message object.

        Prefers ``uniqueBody`` over ``body``, then falls back to
        ``bodyPreview``. Collapses whitespace for easier downstream
        processing.

        Args:
            message: A Graph message resource dictionary.

        Returns:
            Cleaned plain-text body string.
        """
        if not isinstance(message, dict):
            return ""
        unique_body = (message.get("uniqueBody") or {}).get("content", "")
        full_body = (message.get("body") or {}).get("content", "")
        preview = message.get("bodyPreview", "")
        body = unique_body or full_body or preview
        body = body.replace("\r\n", "\n").replace("\r", "\n")
        body = re.sub(r"[ \t]+", " ", body)
        body = re.sub(r"\n{2,}", "\n", body)
        return body.strip()

    # ─── Actions ──────────────────────────────────────────────────────────

    def list_messages(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        folder: str = "inbox",
        subject_filter: Optional[str] = None,
        sender_filter: Optional[str] = None,
        unread_only: bool = False,
        received_after: Optional[str] = None,
        received_before: Optional[str] = None,
        body_keyword: Optional[str] = None,
        top: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        List messages from an Outlook mailbox with optional filtering.

        Filters applied server-side (via OData):
            - subject_filter, sender_filter, unread_only,
              received_after, received_before

        Filters applied client-side (after fetch):
            - body_keyword (substring match against plain-text body)

        Args:
            tenant_id:       Azure AD tenant ID.
            client_id:       App registration client ID.
            client_secret:   App registration client secret.
            mailbox:         Target user mailbox (UPN or object ID).
            folder:          Mail folder name or well-known name (default: "inbox").
            subject_filter:  Return only messages with exactly this subject.
            sender_filter:   Return only messages from this sender address.
            unread_only:     If True, return only unread messages.
            received_after:  ISO 8601 datetime — only messages received at or after.
            received_before: ISO 8601 datetime — only messages received before.
            body_keyword:    Return only messages whose body contains this string.
            top:             Maximum number of messages to fetch (1–1000, default: 25).

        Returns:
            {
                "success": True,
                "count": <int>,
                "messages": [
                    {
                        "id": "...",
                        "subject": "...",
                        "sender": "...",
                        "received_at": "...",
                        "is_read": False,
                        "body_preview": "...",
                        "body": "..."
                    },
                    ...
                ]
            }
        """
        log = self.logger or logging.getLogger(__name__)

        try:
            token = self._get_access_token(tenant_id, client_id, client_secret)
        except requests.HTTPError as exc:
            log.error("[list_messages] Authentication failed: %s", exc)
            return {"success": False, "error": f"Authentication failed: {exc}"}

        try:
            messages_url = self._resolve_folder_url(mailbox, folder, token)
        except ValueError as exc:
            return {"success": False, "error": str(exc)}
        except requests.HTTPError as exc:
            log.error("[list_messages] Folder lookup failed: %s", exc)
            return {"success": False, "error": f"Folder lookup failed: {exc}"}

        # Clamp top to 1–1000
        page_size: int = 25
        if top is not None:
            try:
                page_size = max(1, min(1000, int(top)))
            except (TypeError, ValueError):
                page_size = 25

        odata_filter = self._build_odata_filter(
            subject_filter=subject_filter or None,
            sender_filter=sender_filter or None,
            unread_only=bool(unread_only),
            received_after=received_after or None,
            received_before=received_before or None,
        )

        params: Dict[str, Any] = {
            "$select": "id,sender,subject,receivedDateTime,isRead,body,uniqueBody,bodyPreview",
            "$orderby": "receivedDateTime desc",
            "$top": page_size,
        }
        if odata_filter:
            params["$filter"] = odata_filter

        try:
            data = self._graph_get(messages_url, token, params=params)
        except requests.HTTPError as exc:
            log.error("[list_messages] Graph request failed: %s", exc)
            return {"success": False, "error": f"Graph API error: {exc}"}

        raw_messages: List[Dict[str, Any]] = data.get("value", [])
        log.info("[list_messages] Fetched %d message(s) from '%s/%s'.", len(raw_messages), mailbox, folder)

        results: List[Dict[str, Any]] = []
        for msg in raw_messages:
            body_text = self._extract_body_text(msg)

            # Client-side body keyword filter
            if body_keyword and body_keyword.lower() not in body_text.lower():
                continue

            sender_info = (msg.get("sender") or {}).get("emailAddress", {})
            results.append(
                {
                    "id": msg.get("id", ""),
                    "subject": msg.get("subject", ""),
                    "sender_name": sender_info.get("name", ""),
                    "sender_address": sender_info.get("address", ""),
                    "received_at": msg.get("receivedDateTime", ""),
                    "is_read": msg.get("isRead", None),
                    "body_preview": msg.get("bodyPreview", ""),
                    "body": body_text,
                }
            )

        log.info("[list_messages] Returning %d message(s) after client-side filters.", len(results))
        return {"success": True, "count": len(results), "messages": results}

    # ─────────────────────────────────────────────────────────────────────────

    def get_message(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        message_id: str,
    ) -> Dict[str, Any]:
        """
        Fetch a single message by its Graph message ID.

        Args:
            tenant_id:    Azure AD tenant ID.
            client_id:    App registration client ID.
            client_secret: App registration client secret.
            mailbox:      Target user mailbox (UPN or object ID).
            message_id:   The Graph message ID (``id`` field from list_messages).

        Returns:
            {
                "success": True,
                "message": {
                    "id": "...",
                    "subject": "...",
                    "sender_name": "...",
                    "sender_address": "...",
                    "received_at": "...",
                    "is_read": False,
                    "body_preview": "...",
                    "body": "..."
                }
            }
        """
        log = self.logger or logging.getLogger(__name__)

        if not message_id:
            return {"success": False, "error": "message_id is required."}

        try:
            token = self._get_access_token(tenant_id, client_id, client_secret)
        except requests.HTTPError as exc:
            log.error("[get_message] Authentication failed: %s", exc)
            return {"success": False, "error": f"Authentication failed: {exc}"}

        url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}"
        params = {
            "$select": "id,sender,subject,receivedDateTime,isRead,body,uniqueBody,bodyPreview"
        }

        try:
            msg = self._graph_get(url, token, params=params)
        except requests.HTTPError as exc:
            log.error("[get_message] Graph request failed: %s", exc)
            return {"success": False, "error": f"Graph API error: {exc}"}

        body_text = self._extract_body_text(msg)
        sender_info = (msg.get("sender") or {}).get("emailAddress", {})

        return {
            "success": True,
            "message": {
                "id": msg.get("id", ""),
                "subject": msg.get("subject", ""),
                "sender_name": sender_info.get("name", ""),
                "sender_address": sender_info.get("address", ""),
                "received_at": msg.get("receivedDateTime", ""),
                "is_read": msg.get("isRead", None),
                "body_preview": msg.get("bodyPreview", ""),
                "body": body_text,
            },
        }

    # ─────────────────────────────────────────────────────────────────────────

    def extract_with_regex(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        message_id: str,
        pattern: str,
        flags: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Fetch a message and apply a user-supplied regex pattern to its body.

        This is an optional, modular extraction step. It allows Shuffle
        workflows to perform pattern-based data extraction without tying
        the extraction logic to the message retrieval step.

        Supported flag characters (combinable, e.g. "im"):
            ``i`` — IGNORECASE
            ``m`` — MULTILINE
            ``s`` — DOTALL

        Args:
            tenant_id:    Azure AD tenant ID.
            client_id:    App registration client ID.
            client_secret: App registration client secret.
            mailbox:      Target user mailbox (UPN or object ID).
            message_id:   The Graph message ID to fetch.
            pattern:      Python-compatible regex pattern string.
            flags:        Optional string of flag characters (e.g. "im").

        Returns:
            {
                "success": True,
                "matches": ["match1", "match2", ...],
                "count": <int>
            }
        """
        log = self.logger or logging.getLogger(__name__)

        if not pattern:
            return {"success": False, "error": "pattern is required."}

        # Fetch the message first
        result = self.get_message(tenant_id, client_id, client_secret, mailbox, message_id)
        if not result.get("success"):
            return result

        body = result["message"].get("body", "")

        # Compile regex flags
        re_flags = 0
        for char in (flags or "").lower():
            if char == "i":
                re_flags |= re.IGNORECASE
            elif char == "m":
                re_flags |= re.MULTILINE
            elif char == "s":
                re_flags |= re.DOTALL

        try:
            compiled = re.compile(pattern, re_flags)
            raw_matches = compiled.findall(body)
        except re.error as exc:
            log.error("[extract_with_regex] Invalid regex pattern: %s", exc)
            return {"success": False, "error": f"Invalid regex pattern: {exc}"}

        # Flatten tuple groups (from capturing groups) to strings
        matches: List[str] = []
        for m in raw_matches:
            if isinstance(m, tuple):
                matches.extend(str(g) for g in m if g is not None)
            else:
                matches.append(str(m))

        log.info(
            "[extract_with_regex] Pattern '%s' matched %d result(s) in message '%s'.",
            pattern,
            len(matches),
            message_id,
        )
        return {"success": True, "matches": matches, "count": len(matches)}


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    OutlookMailReader.run()
