#!/usr/bin/env python3
"""CLI entrypoint for the outlook-graph skill."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from auth_manager import AuthConfig, AuthConfigError, AuthError, AuthManager, DependencyError
from graph_client import GraphAPIError, GraphClient

DEFAULT_SELECT_FIELDS = [
    "id",
    "subject",
    "from",
    "receivedDateTime",
    "isRead",
    "hasAttachments",
    "bodyPreview",
    "webLink",
]


warnings.filterwarnings(
    "ignore",
    message="urllib3 v2 only supports OpenSSL",
)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    output_format, cleaned_argv = extract_output_format(argv or sys.argv[1:])

    parser = argparse.ArgumentParser(description="Outlook Graph CLI")

    root = parser.add_subparsers(dest="domain", required=True)

    auth = root.add_parser("auth", help="Authentication operations")
    auth_sub = auth.add_subparsers(dest="action", required=True)

    auth_login = auth_sub.add_parser("login", help="Login with browser or device flow")
    auth_login.add_argument("--method", choices=["browser", "device"], default="browser")
    auth_login.add_argument("--profile", default=None)

    auth_status = auth_sub.add_parser("status", help="Show auth status")
    auth_status.add_argument("--profile", default=None)

    auth_logout = auth_sub.add_parser("logout", help="Clear cached auth for a profile")
    auth_logout.add_argument("--profile", default=None)

    mail = root.add_parser("mail", help="Mail operations")
    mail_sub = mail.add_subparsers(dest="action", required=True)

    mail_list = mail_sub.add_parser("list", help="List messages")
    mail_list.add_argument("--folder", default="inbox")
    mail_list.add_argument("--unread-only", action="store_true")
    mail_list.add_argument("--top", type=int, default=10)
    mail_list.add_argument(
        "--select",
        default=",".join(DEFAULT_SELECT_FIELDS),
        help="Comma-separated Graph message fields",
    )
    mail_list.add_argument("--profile", default=None)

    mail_get = mail_sub.add_parser("get", help="Get one message")
    mail_get.add_argument("--message-id", required=True)
    mail_get.add_argument("--profile", default=None)

    mail_mark = mail_sub.add_parser("mark", help="Mark message read or unread")
    mail_mark.add_argument("--message-id", required=True)
    mail_mark.add_argument("--read", required=True, help="true or false")
    mail_mark.add_argument("--profile", default=None)

    mail_draft = mail_sub.add_parser("draft", help="Create email draft")
    mail_draft.add_argument("--to", nargs="+", required=True, help="Recipient(s)")
    mail_draft.add_argument("--subject", required=True)
    mail_draft.add_argument("--body-file", default=None)
    mail_draft.add_argument("--body", default="")
    mail_draft.add_argument("--body-content-type", choices=["Text", "HTML"], default="Text")
    mail_draft.add_argument("--profile", default=None)

    mail_send = mail_sub.add_parser("send-draft", help="Send an existing draft")
    mail_send.add_argument("--message-id", required=True)
    mail_send.add_argument("--confirm-send", action="store_true")
    mail_send.add_argument("--profile", default=None)

    attachments = root.add_parser("attachments", help="Attachment operations")
    att_sub = attachments.add_subparsers(dest="action", required=True)

    att_list = att_sub.add_parser("list", help="List message attachments")
    att_list.add_argument("--message-id", required=True)
    att_list.add_argument("--profile", default=None)

    att_download = att_sub.add_parser("download", help="Download one attachment")
    att_download.add_argument("--message-id", required=True)
    att_download.add_argument("--attachment-id", required=True)
    att_download.add_argument("--output-dir", default=None)
    att_download.add_argument("--profile", default=None)

    att_download_all = att_sub.add_parser("download-all", help="Download all attachments")
    att_download_all.add_argument("--message-id", required=True)
    att_download_all.add_argument("--output-dir", default=None)
    att_download_all.add_argument("--profile", default=None)

    args = parser.parse_args(cleaned_argv)
    args.format = output_format
    return args


def extract_output_format(argv: List[str]) -> Tuple[str, List[str]]:
    fmt = "json"
    cleaned: List[str] = []
    skip_next = False

    for index, arg in enumerate(argv):
        if skip_next:
            skip_next = False
            continue

        if arg == "--format":
            if index + 1 >= len(argv):
                raise ValueError("--format requires a value: json or text")
            fmt = argv[index + 1].strip().lower()
            skip_next = True
            continue

        if arg.startswith("--format="):
            fmt = arg.split("=", 1)[1].strip().lower()
            continue

        cleaned.append(arg)

    if fmt not in {"json", "text"}:
        raise ValueError("--format must be one of: json, text")

    return fmt, cleaned


def main() -> int:
    output_format = "json"
    try:
        args = parse_args()
        output_format = args.format

        if args.domain == "auth":
            result = run_auth(args)
        elif args.domain == "mail":
            result = run_mail(args)
        elif args.domain == "attachments":
            result = run_attachments(args)
        else:
            raise RuntimeError(f"Unsupported domain '{args.domain}'")

        emit({"ok": True, "result": result}, output_format)
        return 0

    except (AuthConfigError, AuthError, DependencyError, GraphAPIError, ValueError, OSError) as err:
        emit(
            {
                "ok": False,
                "error": {
                    "type": err.__class__.__name__,
                    "message": str(err),
                },
            },
            output_format,
        )
        return 1


def run_auth(args: argparse.Namespace) -> Dict[str, Any]:
    manager = build_auth_manager(args.profile)

    if args.action == "login":
        return manager.login(args.method)
    if args.action == "status":
        return manager.status()
    if args.action == "logout":
        return manager.logout()

    raise RuntimeError(f"Unsupported auth action '{args.action}'")


def run_mail(args: argparse.Namespace) -> Dict[str, Any]:
    if args.action == "list":
        client = build_graph_client(args.profile)
        select_fields = parse_select_fields(args.select)
        messages = client.list_messages(
            folder=args.folder,
            unread_only=args.unread_only,
            top=args.top,
            select_fields=select_fields,
        )
        return {
            "folder": args.folder,
            "unread_only": bool(args.unread_only),
            "top": int(args.top),
            "count": len(messages),
            "messages": messages,
        }

    if args.action == "get":
        client = build_graph_client(args.profile)
        message = client.get_message(args.message_id)
        return {"message": message}

    if args.action == "mark":
        client = build_graph_client(args.profile)
        read_value = parse_boolean(args.read)
        updated = client.mark_message(args.message_id, read=read_value)
        return {
            "message_id": args.message_id,
            "read": read_value,
            "updated": updated,
        }

    if args.action == "draft":
        client = build_graph_client(args.profile)
        recipients = parse_recipients(args.to)
        body = read_body(args.body_file, args.body)
        draft = client.create_draft(
            to_recipients=recipients,
            subject=args.subject,
            body=body,
            body_content_type=args.body_content_type,
        )
        return {
            "message_id": draft.get("id"),
            "subject": draft.get("subject"),
            "to": recipients,
            "is_draft": True,
            "web_link": draft.get("webLink"),
        }

    if args.action == "send-draft":
        if not args.confirm_send:
            raise ValueError(
                "send-draft requires --confirm-send. "
                "Guardrail: create drafts first and require explicit confirmation to send."
            )
        client = build_graph_client(args.profile)
        client.send_draft(args.message_id)
        return {
            "message_id": args.message_id,
            "sent": True,
            "confirmation": "explicit",
        }

    raise RuntimeError(f"Unsupported mail action '{args.action}'")


def run_attachments(args: argparse.Namespace) -> Dict[str, Any]:
    client = build_graph_client(args.profile)

    if args.action == "list":
        attachments = client.list_attachments(args.message_id)
        return {
            "message_id": args.message_id,
            "count": len(attachments),
            "attachments": attachments,
        }

    if args.action == "download":
        target_dir = resolve_output_dir(args.output_dir)
        metadata = client.get_attachment(args.message_id, args.attachment_id)
        payload = download_one_attachment(
            client=client,
            message_id=args.message_id,
            attachment_id=args.attachment_id,
            target_dir=target_dir,
            metadata=metadata,
        )
        return payload

    if args.action == "download-all":
        target_dir = resolve_output_dir(args.output_dir)
        attachments = client.list_attachments(args.message_id)
        downloaded = []
        failed = []

        for attachment in attachments:
            attachment_id = attachment.get("id")
            if not attachment_id:
                failed.append({"reason": "missing id", "attachment": attachment})
                continue
            try:
                result = download_one_attachment(
                    client=client,
                    message_id=args.message_id,
                    attachment_id=attachment_id,
                    target_dir=target_dir,
                    metadata=attachment,
                )
                downloaded.append(result)
            except (GraphAPIError, OSError, ValueError) as err:
                failed.append({
                    "attachment_id": attachment_id,
                    "name": attachment.get("name"),
                    "error": str(err),
                })

        return {
            "message_id": args.message_id,
            "output_dir": str(target_dir),
            "downloaded_count": len(downloaded),
            "failed_count": len(failed),
            "downloaded": downloaded,
            "failed": failed,
        }

    raise RuntimeError(f"Unsupported attachments action '{args.action}'")


def build_auth_manager(profile: Optional[str]) -> AuthManager:
    config = AuthConfig.from_env(profile_override=profile)
    return AuthManager(config)


def build_graph_client(profile: Optional[str]) -> GraphClient:
    manager = build_auth_manager(profile)
    return GraphClient(manager)


def parse_select_fields(raw: str) -> List[str]:
    fields = [field.strip() for field in raw.split(",") if field.strip()]
    return fields or list(DEFAULT_SELECT_FIELDS)


def parse_boolean(raw: str) -> bool:
    normalized = raw.strip().lower()
    if normalized in {"1", "true", "yes", "y"}:
        return True
    if normalized in {"0", "false", "no", "n"}:
        return False
    raise ValueError("--read must be one of: true, false")


def parse_recipients(raw_values: List[str]) -> List[str]:
    recipients: List[str] = []
    for raw in raw_values:
        pieces = re.split(r"[;,]", raw)
        for piece in pieces:
            email = piece.strip()
            if email:
                recipients.append(email)

    if not recipients:
        raise ValueError("At least one recipient is required")

    return recipients


def read_body(body_file: Optional[str], body_inline: str) -> str:
    if body_file:
        path = Path(body_file).expanduser()
        return path.read_text(encoding="utf-8")
    return body_inline or ""


def resolve_output_dir(output_dir: Optional[str]) -> Path:
    raw = output_dir or os.environ.get("OUTLOOK_OUTPUT_DIR") or "./outlook_downloads"
    path = Path(raw).expanduser().resolve()
    path.mkdir(parents=True, exist_ok=True)
    return path


def download_one_attachment(
    client: GraphClient,
    message_id: str,
    attachment_id: str,
    target_dir: Path,
    metadata: Dict[str, Any],
) -> Dict[str, Any]:
    data, content_type = client.download_attachment_bytes(message_id, attachment_id)

    original_name = metadata.get("name") or f"{attachment_id}.bin"
    safe_name = sanitize_filename(str(original_name))
    destination = uniquify_path(target_dir / safe_name)
    destination.write_bytes(data)

    return {
        "attachment_id": attachment_id,
        "name": metadata.get("name"),
        "saved_path": str(destination),
        "size_bytes": len(data),
        "content_type": metadata.get("contentType") or content_type,
        "is_inline": bool(metadata.get("isInline", False)),
    }


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]", "_", name).strip(" .")
    if not cleaned:
        return "attachment.bin"
    if len(cleaned) > 180:
        suffix = Path(cleaned).suffix
        cleaned = cleaned[: 180 - len(suffix)] + suffix
    return cleaned


def uniquify_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for index in range(1, 1000):
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate

    raise OSError(f"Could not find free filename for '{path.name}'")


def emit(payload: Dict[str, Any], fmt: str) -> None:
    if fmt == "json":
        print(json.dumps(payload, indent=2, sort_keys=True))
        return

    if payload.get("ok"):
        result = payload.get("result", {})
        print(render_text(result))
    else:
        err = payload.get("error", {})
        print(f"ERROR [{err.get('type')}]: {err.get('message')}")


def render_text(result: Any) -> str:
    if isinstance(result, (str, int, float, bool)) or result is None:
        return str(result)

    if isinstance(result, list):
        return "\n".join(f"- {json.dumps(item, sort_keys=True)}" for item in result)

    if isinstance(result, dict):
        lines: List[str] = []
        for key in sorted(result.keys()):
            value = result[key]
            if isinstance(value, (dict, list)):
                lines.append(f"{key}:")
                lines.append(json.dumps(value, indent=2, sort_keys=True))
            else:
                lines.append(f"{key}: {value}")
        return "\n".join(lines)

    return json.dumps(result, sort_keys=True)


if __name__ == "__main__":
    raise SystemExit(main())
