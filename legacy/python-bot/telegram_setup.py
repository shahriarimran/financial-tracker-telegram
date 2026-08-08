from __future__ import annotations

import os
import sys

import requests
from dotenv import load_dotenv


def main() -> int:
    load_dotenv()

    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    if not token:
        print("Set TELEGRAM_BOT_TOKEN in .env first.")
        return 1

    base = f"https://api.telegram.org/bot{token}"

    me = requests.get(f"{base}/getMe", timeout=20)
    me.raise_for_status()
    me_json = me.json()
    if not me_json.get("ok"):
        print("Bot token is invalid:", me_json)
        return 1

    bot = me_json["result"]
    print(f"Bot OK: @{bot.get('username', '<no username>')}")
    print("Send /start to the bot in Telegram, then run this script again.\n")

    updates = requests.get(f"{base}/getUpdates", timeout=20)
    updates.raise_for_status()
    payload = updates.json()

    chats = {}
    for update in payload.get("result", []):
        message = (
            update.get("message")
            or update.get("edited_message")
            or update.get("channel_post")
            or update.get("edited_channel_post")
        )
        if not message:
            continue

        chat = message.get("chat", {})
        chat_id = chat.get("id")
        if chat_id is None:
            continue

        label = (
            chat.get("title")
            or chat.get("username")
            or " ".join(
                part
                for part in [chat.get("first_name"), chat.get("last_name")]
                if part
            )
            or "Unnamed chat"
        )
        chats[str(chat_id)] = label

    if not chats:
        print("No chats found yet. Send /start to the bot and rerun.")
        return 0

    print("Available chat IDs:")
    for chat_id, label in chats.items():
        print(f"  {chat_id}  ({label})")

    print("\nCopy the desired ID into TELEGRAM_CHAT_ID in .env.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
