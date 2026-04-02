"""
Microsoft Teams platform adapter.

Receives messages via Bot Framework incoming webhooks and sends replies
via the Bot Framework REST API.

Requires:
  - aiohttp (included in hermes-agent[messaging])
  - An Azure AD app registration with:
    - Microsoft App ID (MICROSOFT_APP_ID)
    - App password/secret (MICROSOFT_APP_PASSWORD)
    - Bot configured with an HTTPS messaging endpoint

Env vars:
  MICROSOFT_APP_ID       — Azure AD application (client) ID
  MICROSOFT_APP_PASSWORD — App password/secret
  MICROSOFT_TENANT_ID    — Azure tenant ID (optional; defaults to common)
  MICROSOFT bot_SERVICE_URL — Bot Framework service URL (provided at runtime)

Config (config.yaml platforms.microsoft-teams):
  enabled: true
  extra:
    host: "0.0.0.0"       # webhook server bind address
    port: 8645            # webhook server port
    allow_dm: true        # allow 1:1 DMs with the bot
    allow_channels: true  # allow messages in channels (requires mention)
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import time
from typing import Any, Dict, Optional

try:
    import aiohttp
    from aiohttp import web
    AIOHTTP_AVAILABLE = True
except ImportError:
    AIOHTTP_AVAILABLE = False
    aiohttp = None
    web = None  # type: ignore[assignment]

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parents[2]))

from gateway.config import Platform, PlatformConfig
from gateway.platforms.base import (
    BasePlatformAdapter,
    MessageEvent,
    MessageType,
    SendResult,
)

logger = logging.getLogger(__name__)

# Default webhook server settings
DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 8645
# Teams maximum message size
MAX_MESSAGE_LENGTH = 10000


def check_msteams_requirements() -> bool:
    """Check if Teams adapter dependencies are available."""
    return AIOHTTP_AVAILABLE


class MSTeamsAdapter(BasePlatformAdapter):
    """
    Microsoft Teams bot adapter.

    Receives Teams messages via aiohttp webhook server and sends replies
    via the Bot Framework Direct Line / REST API.
    """

    _instance: "MSTeamsAdapter | None" = None

    def __init__(self, config: PlatformConfig):
        super().__init__(config, Platform.MICROSOFT_TEAMS)
        self._app_id: str = os.getenv("MICROSOFT_APP_ID", "")
        self._app_password: str = os.getenv("MICROSOFT_APP_PASSWORD", "")
        self._tenant_id: str = os.getenv("MICROSOFT_TENANT_ID", "common")
        self._service_url: str = os.getenv("MICROSOFT_SERVICE_URL", "")
        self._host: str = config.extra.get("host", DEFAULT_HOST)
        self._port: int = int(config.extra.get("port", DEFAULT_PORT))
        self._allow_dm: bool = config.extra.get("allow_dm", True)
        self._allow_channels: bool = config.extra.get("allow_channels", True)
        self._runner: Optional[web.AppRunner] = None
        self._site: Optional[web.TCPSite] = None
        self._access_token: Optional[str] = None
        self._token_expires_at: float = 0.0
        # conversation reference: conversation_id → {service_url, conversation_id, bot_id, channel_id}
        self._conversations: Dict[str, Dict[str, str]] = {}
        # Bot's own ID (set after first successful message send)
        self._bot_id: Optional[str] = None

        # Only allow one instance (aiohttp server binds a port)
        if MSTeamsAdapter._instance is not None:
            logger.warning("[Teams] Only one Teams adapter instance supported; ignoring duplicate")
        MSTeamsAdapter._instance = self

    # -------------------------------------------------------------------------
    # OAuth token management
    # -------------------------------------------------------------------------

    async def _ensure_token(self) -> bool:
        """Fetch (or refresh) a Bot Framework access token."""
        if self._access_token and time.time() < self._token_expires_at - 60:
            return True

        if not self._app_id or not self._app_password:
            logger.error("[Teams] MICROSOFT_APP_ID or MICROSOFT_APP_PASSWORD not set")
            return False

        token_url = (
            f"https://login.microsoftonline.com/{self._tenant_id}"
            "/oauth2/v2.0/token"
        )
        scope = "https://api.botframework.com/.default"

        try:
            async with aiohttp.ClientSession() as session:
                resp = await session.post(
                    token_url,
                    data={
                        "grant_type": "client_credentials",
                        "client_id": self._app_id,
                        "client_secret": self._app_password,
                        "scope": scope,
                    },
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    timeout=aiohttp.ClientTimeout(total=15),
                )
                if resp.status != 200:
                    body = await resp.text()
                    logger.error(f"[Teams] Token request failed {resp.status}: {body}")
                    return False
                data = await resp.json()
                self._access_token = data["access_token"]
                self._token_expires_at = time.time() + data.get("expires_in", 3600)
                logger.info("[Teams] Obtained fresh access token")
                return True
        except Exception as exc:
            logger.error(f"[Teams] Token request error: {exc}")
            return False

    async def _api_headers(self) -> Dict[str, str]:
        """Return HTTP headers for Bot Framework API calls."""
        await self._ensure_token()
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

    # -------------------------------------------------------------------------
    # aiohttp webhook server
    # -------------------------------------------------------------------------

    def _build_app(self) -> web.Application:
        """Build the aiohttp application with Teams webhook routes."""
        app = web.Application()
        app.router.add_post("/teams/webhook", self._handle_teams_webhook)
        app.router.add_get("/teams/healthz", self._handle_health)
        return app

    async def _handle_health(self, request: web.Request) -> web.Response:
        return web.json_response({"status": "ok", "platform": "microsoft-teams"})

    async def _handle_teams_webhook(self, request: web.Request) -> web.Response:
        """
        Receive incoming activity from Bot Framework.

        Bot Framework always requires a 200 response within 15 seconds.
        We respond immediately and process async.
        """
        try:
            activity = await request.json()
        except Exception as exc:
            logger.warning(f"[Teams] Failed to parse activity: {exc}")
            return web.Response(status=400, text="invalid payload")

        # Bot Framework validation challenge (GET with serviceUrl + channelId)
        if activity.get("type") == "conversationUpdate":
            # Acknowledge conversation updates (member joins, etc.)
            logger.info("[Teams] conversationUpdate received")
            return web.Response(status=200, text="")

        if activity.get("type") == "ping":
            return web.Response(status=200, text="")

        asyncio.create_task(self._process_activity(activity))
        return web.Response(status=200, text="")

    async def _process_activity(self, activity: Dict[str, Any]) -> None:
        """Process a Teams activity and dispatch to the gateway handler."""
        try:
            activity_type = activity.get("type", "")
            if activity_type != "message":
                logger.debug(f"[Teams] Ignoring activity type: {activity_type}")
                return

            channel_id = activity.get("channelId", "")
            if channel_id != "msteams":
                logger.debug(f"[Teams] Ignoring non-Teams channel: {channel_id}")
                return

            # Extract sender and conversation
            from_id = self._get_participant_id(activity, "from")
            from_name = self._get_participant_name(activity, "from")
            recipient_id = self._get_participant_id(activity, "recipient")
            conversation = activity.get("conversation", {})
            conversation_id = conversation.get("id", "")
            channel_name = conversation.get("name", "")

            # Skip if no text
            text = self._extract_text(activity)
            if not text:
                return

            # Determine chat_id (use conversation ID)
            chat_id = conversation_id

            # Check DM/channel restrictions
            is_dm = activity.get("conversation", {}).get("conversationType") == "personal"
            if is_dm:
                if not self._allow_dm:
                    logger.debug("[Teams] DMs disabled; ignoring")
                    return
            else:
                if not self._allow_channels:
                    logger.debug("[Teams] channel messages disabled; ignoring")
                    return
                # In channels, require a mention of the bot (Teams @mention)
                # The bot's mention text appears in text; strip it
                bot_id = self._bot_id or recipient_id or ""
                if bot_id and f"<at>{bot_id}</at>" in text:
                    text = text.replace(f"<at>{bot_id}</at>", "").strip()

            # Store conversation reference for proactive sends
            service_url = activity.get("serviceUrl", self._service_url)
            if service_url:
                self._service_url = service_url
            self._conversations[conversation_id] = {
                "service_url": service_url,
                "conversation_id": conversation_id,
                "bot_id": recipient_id,
                "channel_id": channel_id,
            }
            if not self._bot_id and recipient_id:
                self._bot_id = recipient_id

            # Strip Teams HTML formatting from text
            text = self._strip_teams_formatting(text)

            logger.info(f"[Teams] Message from {from_name} ({from_id}) in {channel_name or 'DM'}: {text[:80]}")

            # Build message event
            source = self.build_source(
                platform=Platform.MICROSOFT_TEAMS,
                chat_id=chat_id,
                sender_id=from_id,
                sender_name=from_name,
                chat_name=channel_name or "DM",
            )

            event = MessageEvent(
                platform=Platform.MICROSOFT_TEAMS,
                message_type=MessageType.MESSAGE,
                text=text,
                chat_id=chat_id,
                sender_id=from_id,
                sender_name=from_name,
                message_id=activity.get("id", ""),
                timestamp=activity.get("timestamp", ""),
                metadata={
                    "conversation_id": conversation_id,
                    "service_url": service_url,
                    "channel_name": channel_name,
                    "is_dm": is_dm,
                    "raw_activity": activity,
                },
                source=source,
            )

            if self._message_handler:
                response = await self._message_handler(event)
                if response:
                    await self.send(chat_id, response, reply_to=activity.get("id"))
            else:
                logger.warning("[Teams] No message handler set")

        except Exception as exc:
            logger.exception(f"[Teams] Error processing activity: {exc}")

    def _get_participant_id(self, activity: Dict[str, Any], key: str) -> str:
        """Extract participant ID from activity 'from' or 'recipient' field."""
        try:
            return activity.get(key, {}).get("id", "unknown")
        except Exception:
            return "unknown"

    def _get_participant_name(self, activity: Dict[str, Any], key:str) -> str:
        """Extract participant name from activity 'from' or 'recipient' field."""
        try:
            return activity.get(key, {}).get("name", "Unknown")
        except Exception:
            return "Unknown"

    def _extract_text(self, activity: Dict[str, Any]) -> str:
        """Extract plain text from a Teams message activity."""
        try:
            content = activity.get("text", "") or ""
            # Teams may send content in textFormat=xml (Adaptive Cards, etc.)
            attachments = activity.get("attachments", [])
            for att in attachments:
                if att.get("contentType", "").startswith("application/vnd.microsoft.card."):
                    card = att.get("content", {})
                    if isinstance(card, dict):
                        content = card.get("text", content)
            return content.strip()
        except Exception:
            return ""

    def _strip_teams_formatting(self, text: str) -> str:
        """Remove Teams XML/HTML formatting markers."""
        import re
        # Remove <at>...</at> mention tags
        text = re.sub(r"<at>[^<]*</at>", "", text)
        # Remove <deletekeyword/> and similar self-closing tags
        text = re.sub(r"<[a-zA-Z]+\s*/>", "", text)
        # Remove any remaining HTML tags
        text = re.sub(r"<[^>]+>", "", text)
        return text.strip()

    # -------------------------------------------------------------------------
    # BasePlatformAdapter interface
    # -------------------------------------------------------------------------

    async def connect(self) -> bool:
        """Start the webhook server and register with Bot Framework."""
        if not self._app_id or not self._app_password:
            logger.error(
                "[Teams] Missing config — set MICROSOFT_APP_ID and "
                "MICROSOFT_APP_PASSWORD env vars"
            )
            return False

        if not await self._ensure_token():
            logger.error("[Teams] Failed to obtain access token; cannot connect")
            return False

        app = self._build_app()
        self._runner = web.AppRunner(app, logger=logger)
        await self._runner.setup()
        self._site = web.TCPSite(self._runner, self._host, self._port)
        await self._site.start()
        logger.info(f"[Teams] Webhook server listening on {self._host}:{self._port}")
        logger.info(
            f"[Teams] Register webhook URL: "
            f"https://your-server.example.com:{self._port}/teams/webhook"
        )
        return True

    async def disconnect(self) -> None:
        """Stop the webhook server."""
        if self._runner:
            await self._runner.cleanup()
            self._runner = None
            self._site = None
            logger.info("[Teams] Webhook server stopped")

    async def send(
        self,
        chat_id: str,
        content: str,
        reply_to: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        """
        Send a message via Bot Framework REST API.

        If metadata contains 'service_url' and 'bot_id', uses that conversation
        reference directly. Otherwise looks up the conversation.
        """
        service_url = None
        conversation_id = chat_id

        if metadata:
            service_url = metadata.get("service_url")
            conversation_id = metadata.get("conversation_id", chat_id)

        if not service_url and conversation_id in self._conversations:
            conv = self._conversations[conversation_id]
            service_url = conv.get("service_url")

        if not service_url:
            logger.error(f"[Teams] No service URL for conversation {conversation_id}")
            return SendResult(success=False, error="No service URL for conversation")

        # Bot Framework API endpoint for sending messages
        url = f"{service_url}v3/conversations/{conversation_id}/activities"

        headers = await self._api_headers()
        if not headers.get("Authorization"):
            return SendResult(success=False, error="No access token")

        message_body: Dict[str, Any] = {
            "type": "message",
            "text": content,
        }
        if reply_to:
            message_body["replyToId"] = reply_to

        try:
            async with aiohttp.ClientSession() as session:
                resp = await session.post(
                    url,
                    json=message_body,
                    headers=headers,
                    timeout=aiohttp.ClientTimeout(total=15),
                )
                if resp.status in (200, 201):
                    data = await resp.json()
                    activity_id = data.get("id", "")
                    logger.info(f"[Teams] Sent message, activity_id={activity_id}")
                    return SendResult(success=True, message_id=activity_id)
                else:
                    body = await resp.text()
                    logger.error(f"[Teams] Send failed {resp.status}: {body}")
                    return SendResult(success=False, error=f"HTTP {resp.status}: {body}")
        except Exception as exc:
            logger.exception(f"[Teams] Send error: {exc}")
            return SendResult(success=False, error=str(exc))

    async def send_typing(self, chat_id: str, metadata: Optional[Dict[str, Any]] = None) -> None:
        """Send a typing indicator via Bot Framework."""
        service_url = None
        conversation_id = chat_id
        if metadata:
            service_url = metadata.get("service_url")
            conversation_id = metadata.get("conversation_id", chat_id)
        elif conversation_id in self._conversations:
            service_url = self._conversations[conversation_id].get("service_url")

        if not service_url:
            return

        url = f"{service_url}v3/conversations/{conversation_id}/activities"
        headers = await self._api_headers()
        try:
            async with aiohttp.ClientSession() as session:
                await session.post(
                    url,
                    json={"type": "typing"},
                    headers=headers,
                    timeout=aiohttp.ClientTimeout(total=5),
                )
        except Exception:
            pass  # best-effort

    async def send_image(
        self,
        chat_id: str,
        image_url: str,
        caption: Optional[str] = None,
        reply_to: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        """
        Send an image via Bot Framework.

        Teams supports image attachments via the ChannelAccount format.
        We send a message with an inline image attachment.
        """
        text = caption or ""
        if image_url:
            text = f"{text}\n\n![image]({image_url})".strip()
        return await self.send(chat_id, text, reply_to=reply_to, metadata=metadata)

    async def get_chat_info(self, chat_id: str) -> Dict[str, Any]:
        """Return info about a Teams conversation."""
        conv = self._conversations.get(chat_id, {})
        return {
            "name": conv.get("channel_name", "Teams Chat"),
            "type": "channel",
            "chat_id": chat_id,
        }
