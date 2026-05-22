#!/usr/bin/env python3
"""Outlook Calendar MCP server for personal Microsoft accounts (device code auth)."""
import asyncio
import json
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path

import msal
import requests
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "")
SCOPES = ["Calendars.ReadWrite"]
TOKEN_CACHE_PATH = Path.home() / ".outlook_mcp_tokens.json"
FLOW_CACHE_PATH = Path.home() / ".outlook_mcp_flow.json"

_app = None


def _get_app():
    global _app
    if _app is None:
        cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_PATH.exists():
            cache.deserialize(TOKEN_CACHE_PATH.read_text())
        _app = msal.PublicClientApplication(
            CLIENT_ID,
            authority="https://login.microsoftonline.com/consumers",
            token_cache=cache,
        )
    return _app


def _save_cache():
    if _app and _app.token_cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(_app.token_cache.serialize())
        TOKEN_CACHE_PATH.chmod(0o600)


def _get_token():
    a = _get_app()
    accounts = a.get_accounts()
    if accounts:
        r = a.acquire_token_silent(SCOPES, account=accounts[0])
        if r and "access_token" in r:
            _save_cache()
            return r["access_token"], None
    return None, "未認証。start_auth → ブラウザ認証 → complete_auth の順で実行してください。"


server = Server("outlook-mcp")


@server.list_tools()
async def list_tools():
    return [
        types.Tool(
            name="start_auth",
            description="Microsoft認証を開始。URLとコードを返すので、ブラウザでサインインする",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="complete_auth",
            description="start_auth後にブラウザ認証が完了したら呼ぶ。トークンを保存する",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="get_calendar_events",
            description="カレンダーの予定を取得",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {"type": "integer", "default": 7, "description": "何日先まで取得するか"}
                },
            },
        ),
        types.Tool(
            name="create_calendar_event",
            description="カレンダーに予定を追加",
            inputSchema={
                "type": "object",
                "required": ["subject", "start", "end"],
                "properties": {
                    "subject": {"type": "string", "description": "タイトル"},
                    "start": {"type": "string", "description": "開始時刻 例: 2026-05-20T15:00:00"},
                    "end": {"type": "string", "description": "終了時刻 例: 2026-05-20T16:00:00"},
                    "body": {"type": "string", "description": "本文（省略可）"},
                },
            },
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict):
    if name == "start_auth":
        a = _get_app()
        accounts = a.get_accounts()
        if accounts:
            return [types.TextContent(type="text", text=f"認証済み: {accounts[0]['username']}")]

        def _start():
            flow = a.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                return None, flow.get("error_description", "デバイスフロー開始失敗")
            FLOW_CACHE_PATH.write_text(json.dumps(flow))
            FLOW_CACHE_PATH.chmod(0o600)
            return flow["message"], None

        msg, err = await asyncio.to_thread(_start)
        if err:
            return [types.TextContent(type="text", text=f"エラー: {err}")]
        return [types.TextContent(type="text", text=msg)]

    if name == "complete_auth":
        if not FLOW_CACHE_PATH.exists():
            return [types.TextContent(type="text", text="start_auth を先に実行してください。")]

        def _complete():
            flow = json.loads(FLOW_CACHE_PATH.read_text())
            a = _get_app()
            result = a.acquire_token_by_device_flow(flow)
            FLOW_CACHE_PATH.unlink(missing_ok=True)
            if "access_token" in result:
                _save_cache()
                accounts = a.get_accounts()
                user = accounts[0]["username"] if accounts else "不明"
                return f"認証成功: {user}"
            return f"失敗: {result.get('error_description', str(result))}"

        msg = await asyncio.to_thread(_complete)
        return [types.TextContent(type="text", text=msg)]

    # 以下は認証済み前提
    token, err = _get_token()
    if err:
        return [types.TextContent(type="text", text=err)]

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    if name == "get_calendar_events":
        days = max(1, min(int(arguments.get("days", 7)), 365))
        now = datetime.now(timezone.utc)
        end_dt = now + timedelta(days=days)

        def _fetch():
            start_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
            end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            url = (
                "https://graph.microsoft.com/v1.0/me/calendarView"
                f"?startDateTime={start_str}&endDateTime={end_str}"
                "&$select=subject,start,end&$orderby=start/dateTime&$top=20"
            )
            return requests.get(url, headers=headers, timeout=30).json()

        data = await asyncio.to_thread(_fetch)
        if "value" not in data:
            return [types.TextContent(type="text", text=f"Error: {data}")]
        events = data["value"]
        if not events:
            return [types.TextContent(type="text", text="予定なし")]
        jst = timezone(timedelta(hours=9))
        lines = []
        for e in events:
            raw = e["start"]["dateTime"]
            tz_name = e["start"].get("timeZone", "UTC")
            dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            dt_jst = dt.astimezone(jst)
            dt_str = dt_jst.strftime("%Y-%m-%d %H:%M")
            lines.append(f"• {dt_str}  {e['subject']}")
        return [types.TextContent(type="text", text="\n".join(lines))]

    if name == "create_calendar_event":
        payload = {
            "subject": arguments["subject"],
            "start": {"dateTime": arguments["start"], "timeZone": "Asia/Tokyo"},
            "end": {"dateTime": arguments["end"], "timeZone": "Asia/Tokyo"},
        }
        if "body" in arguments:
            payload["body"] = {"contentType": "text", "content": arguments["body"]}

        def _create():
            return requests.post(
                "https://graph.microsoft.com/v1.0/me/events",
                headers=headers,
                json=payload,
                timeout=30,
            )

        resp = await asyncio.to_thread(_create)
        if resp.status_code == 201:
            return [types.TextContent(type="text", text=f"追加しました: {arguments['subject']}")]
        return [types.TextContent(type="text", text=f"Error {resp.status_code}: {resp.text}")]

    return [types.TextContent(type="text", text=f"Unknown tool: {name}")]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
