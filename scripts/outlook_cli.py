#!/usr/bin/env python3
"""CLI entrypoint for the outlook-graph skill."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shlex
import shutil
import sys
import uuid
import warnings
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

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
DEFAULT_FOLDER_FIELDS = [
    "id",
    "display_name",
    "parent_id",
    "child_folder_count",
    "total_item_count",
    "unread_item_count",
    "is_hidden",
]
DEFAULT_AUTH_SCOPES = ["User.Read", "Mail.ReadWrite", "Mail.Send"]
RESERVED_OIDC_SCOPES = {"openid", "profile", "offline_access"}
DEFAULT_FOLDER_ROOT = "inbox"
DEFAULT_MAX_FOLDER_NODES = 5000
DEFAULT_RECENT_TOP = 10
DEFAULT_OVERLAP_HOURS = 48
DEFAULT_MAX_PAGES = 20
DEFAULT_MAX_MESSAGES = 1000
FIRST_RUN_BACKFILL_DAYS = 15
STATE_VERSION = 1


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

    auth_onboard = auth_sub.add_parser(
        "onboard",
        help="Return deterministic onboarding steps for agent-driven setup",
    )
    auth_onboard.add_argument("--profile", default=None)
    auth_onboard.add_argument("--client-id", default=None)
    auth_onboard.add_argument("--tenant-id", default=None)
    auth_onboard.add_argument("--redirect-uri", default=None)
    auth_onboard.add_argument("--scopes", default=None)
    auth_onboard.add_argument("--method", choices=["browser", "device"], default="browser")

    mail = root.add_parser("mail", help="Mail operations")
    mail_sub = mail.add_subparsers(dest="action", required=True)

    mail_list = mail_sub.add_parser("list", help="List messages")
    _add_folder_selector_args(mail_list)
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

    folders = root.add_parser("folders", help="Folder discovery operations")
    folders_sub = folders.add_subparsers(dest="action", required=True)

    folders_tree = folders_sub.add_parser("tree", help="Return folder tree and flat index")
    folders_tree.add_argument("--root", default=DEFAULT_FOLDER_ROOT)
    folders_tree.add_argument("--include-hidden", action="store_true")
    folders_tree.add_argument("--max-nodes", type=int, default=DEFAULT_MAX_FOLDER_NODES)
    folders_tree.add_argument("--profile", default=None)

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

    att_recent = att_sub.add_parser(
        "download-recent",
        help="Download attachments from the latest emails in a folder",
    )
    _add_folder_selector_args(att_recent)
    att_recent.add_argument("--top", type=int, default=DEFAULT_RECENT_TOP)
    att_recent.add_argument("--unread-only", action="store_true")
    att_recent.add_argument("--output-dir", default=None)
    att_recent.add_argument("--force-redownload", action="store_true")
    att_recent.add_argument("--profile", default=None)

    att_new = att_sub.add_parser(
        "download-new",
        help="Download only attachments not yet downloaded according to local ledger state",
    )
    _add_folder_selector_args(att_new)
    att_new.add_argument("--output-dir", default=None)
    att_new.add_argument("--overlap-hours", type=int, default=DEFAULT_OVERLAP_HOURS)
    att_new.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES)
    att_new.add_argument("--max-messages", type=int, default=DEFAULT_MAX_MESSAGES)
    att_new.add_argument("--profile", default=None)

    att_state = att_sub.add_parser("state", help="Inspect or reset attachment download state")
    state_sub = att_state.add_subparsers(dest="state_action", required=True)

    state_status = state_sub.add_parser("status", help="Show state for one folder stream")
    _add_folder_selector_args(state_status)
    state_status.add_argument("--profile", default=None)

    state_reset = state_sub.add_parser("reset", help="Reset state for one folder stream")
    _add_folder_selector_args(state_reset)
    state_reset.add_argument("--confirm-reset", action="store_true")
    state_reset.add_argument("--profile", default=None)

    args = parser.parse_args(cleaned_argv)
    args.format = output_format
    return args


def _add_folder_selector_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--folder", default=DEFAULT_FOLDER_ROOT)
    parser.add_argument("--folder-id", default=None)
    parser.add_argument("--folder-path", default=None)
    parser.add_argument("--include-hidden-folders", action="store_true")


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
        elif args.domain == "folders":
            result = run_folders(args)
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
    if args.action == "onboard":
        return build_onboarding_plan(args)

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
        if args.top <= 0:
            raise ValueError("--top must be greater than 0")
        client = build_graph_client(args.profile)
        select_fields = parse_select_fields(args.select)
        folder_resolution = resolve_folder_selector(
            client=client,
            folder=args.folder,
            folder_id=args.folder_id,
            folder_path=args.folder_path,
            include_hidden=bool(args.include_hidden_folders),
        )
        messages = client.list_messages(
            folder=folder_resolution["folder_token"],
            unread_only=bool(args.unread_only),
            top=args.top,
            select_fields=select_fields,
        )
        return {
            "folder": args.folder,
            "folder_mode": folder_resolution["folder_mode"],
            "folder_input": folder_resolution["folder_input"],
            "resolved_folder_id": folder_resolution["resolved_folder_id"],
            "resolved_folder_path": folder_resolution.get("resolved_folder_path"),
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


def run_folders(args: argparse.Namespace) -> Dict[str, Any]:
    if args.action != "tree":
        raise RuntimeError(f"Unsupported folders action '{args.action}'")

    if args.max_nodes <= 0:
        raise ValueError("--max-nodes must be greater than 0")

    client = build_graph_client(args.profile)
    root_token = (args.root or DEFAULT_FOLDER_ROOT).strip() or DEFAULT_FOLDER_ROOT
    tree_payload = build_folder_tree(
        client=client,
        root_token=root_token,
        include_hidden=bool(args.include_hidden),
        max_nodes=int(args.max_nodes),
    )

    return {
        "root_token": root_token,
        "include_hidden": bool(args.include_hidden),
        "max_nodes": int(args.max_nodes),
        "count": len(tree_payload["index"]),
        "tree": tree_payload["tree"],
        "index": tree_payload["index"],
    }


def run_attachments(args: argparse.Namespace) -> Dict[str, Any]:
    if args.action == "state":
        return run_attachment_state(args)

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

    if args.action == "download-recent":
        return run_attachments_download_recent(client, args)

    if args.action == "download-new":
        return run_attachments_download_new(client, args)

    raise RuntimeError(f"Unsupported attachments action '{args.action}'")


def run_attachment_state(args: argparse.Namespace) -> Dict[str, Any]:
    client = build_graph_client(args.profile)
    profile = resolve_profile(args.profile)

    resolution = resolve_folder_selector(
        client=client,
        folder=args.folder,
        folder_id=args.folder_id,
        folder_path=args.folder_path,
        include_hidden=bool(args.include_hidden_folders),
    )
    account_home_id = get_authenticated_account_home_id(client.auth)
    stream_paths = build_stream_paths(
        profile=profile,
        account_home_id=account_home_id,
        folder_id=str(resolution["resolved_folder_id"] or resolution["folder_token"]),
    )

    if args.state_action == "status":
        state = load_stream_state(
            stream_paths=stream_paths,
            profile=profile,
            account_home_id=account_home_id,
            folder_id=str(resolution["resolved_folder_id"] or resolution["folder_token"]),
            folder_path=resolution.get("resolved_folder_path"),
        )
        completed_keys = load_completed_keys(stream_paths)
        exists = stream_paths["state_file"].exists() or stream_paths["ledger_file"].exists()
        return {
            "exists": exists,
            "stream_id": stream_paths["stream_id"],
            "state_dir": str(stream_paths["stream_dir"]),
            "state_file": str(stream_paths["state_file"]),
            "ledger_file": str(stream_paths["ledger_file"]),
            "folder_mode": resolution["folder_mode"],
            "folder_input": resolution["folder_input"],
            "resolved_folder_id": resolution["resolved_folder_id"],
            "resolved_folder_path": resolution.get("resolved_folder_path"),
            "state": state,
            "pending_failures_count": len(state.get("pending_failures", {})),
            "completed_keys_count": len(completed_keys),
            "ledger_entries": count_ledger_entries(stream_paths),
        }

    if args.state_action == "reset":
        if not args.confirm_reset:
            raise ValueError("state reset requires --confirm-reset")
        removed = False
        if stream_paths["stream_dir"].exists():
            shutil.rmtree(stream_paths["stream_dir"])
            removed = True
        return {
            "removed": removed,
            "stream_id": stream_paths["stream_id"],
            "state_dir": str(stream_paths["stream_dir"]),
            "folder_mode": resolution["folder_mode"],
            "folder_input": resolution["folder_input"],
            "resolved_folder_id": resolution["resolved_folder_id"],
            "resolved_folder_path": resolution.get("resolved_folder_path"),
        }

    raise RuntimeError(f"Unsupported attachments state action '{args.state_action}'")


def run_attachments_download_recent(client: GraphClient, args: argparse.Namespace) -> Dict[str, Any]:
    if args.top <= 0:
        raise ValueError("--top must be greater than 0")

    profile = resolve_profile(args.profile)
    resolution = resolve_folder_selector(
        client=client,
        folder=args.folder,
        folder_id=args.folder_id,
        folder_path=args.folder_path,
        include_hidden=bool(args.include_hidden_folders),
    )
    account_home_id = get_authenticated_account_home_id(client.auth)
    stream_paths = build_stream_paths(
        profile=profile,
        account_home_id=account_home_id,
        folder_id=str(resolution["resolved_folder_id"] or resolution["folder_token"]),
    )
    state = load_stream_state(
        stream_paths=stream_paths,
        profile=profile,
        account_home_id=account_home_id,
        folder_id=str(resolution["resolved_folder_id"] or resolution["folder_token"]),
        folder_path=resolution.get("resolved_folder_path"),
    )
    completed_keys = load_completed_keys(stream_paths)

    messages = client.list_messages(
        folder=resolution["folder_token"],
        unread_only=bool(args.unread_only),
        top=int(args.top),
        select_fields=list(DEFAULT_SELECT_FIELDS),
        has_attachments_only=True,
    )

    run_id = uuid.uuid4().hex
    run_started = now_utc()
    run_started_iso = iso_utc(run_started)

    output_dir = resolve_output_dir(args.output_dir)
    batch_dir = create_batch_output_dir(output_dir, "recent")

    append_ledger_event(
        stream_paths,
        {
            "event": "run_started",
            "run_id": run_id,
            "mode": "download_recent",
            "folder_id": resolution["resolved_folder_id"],
            "folder_path": resolution.get("resolved_folder_path"),
            "top": int(args.top),
            "force_redownload": bool(args.force_redownload),
            "started_at": run_started_iso,
        },
    )

    message_results = []
    downloaded_total = 0
    skipped_total = 0
    failed_total = 0

    for message in messages:
        result = process_message_attachments(
            client=client,
            message=message,
            batch_dir=batch_dir,
            completed_keys=completed_keys,
            state=state,
            stream_paths=stream_paths,
            run_id=run_id,
            force_redownload=bool(args.force_redownload),
        )
        message_results.append(result)
        downloaded_total += result["downloaded_count"]
        skipped_total += result["skipped_count"]
        failed_total += result["failed_count"]

    run_completed = now_utc()
    run_completed_iso = iso_utc(run_completed)

    state["profile"] = profile
    state["account_home_id"] = account_home_id
    state["folder_id"] = str(resolution["resolved_folder_id"] or resolution["folder_token"])
    state["folder_path"] = resolution.get("resolved_folder_path")
    state["last_run_started_utc"] = run_started_iso
    state["last_run_completed_utc"] = run_completed_iso
    state["downloaded_total"] = int(state.get("downloaded_total", 0)) + downloaded_total
    state["skipped_total"] = int(state.get("skipped_total", 0)) + skipped_total
    state["failed_total"] = int(state.get("failed_total", 0)) + failed_total

    save_stream_state(stream_paths, state)

    append_ledger_event(
        stream_paths,
        {
            "event": "run_completed",
            "run_id": run_id,
            "mode": "download_recent",
            "started_at": run_started_iso,
            "completed_at": run_completed_iso,
            "messages_scanned": len(messages),
            "downloaded_count": downloaded_total,
            "skipped_count": skipped_total,
            "failed_count": failed_total,
        },
    )

    return {
        "run_id": run_id,
        "mode": "download_recent",
        "folder_mode": resolution["folder_mode"],
        "folder_input": resolution["folder_input"],
        "resolved_folder_id": resolution["resolved_folder_id"],
        "resolved_folder_path": resolution.get("resolved_folder_path"),
        "top": int(args.top),
        "unread_only": bool(args.unread_only),
        "force_redownload": bool(args.force_redownload),
        "output_dir": str(batch_dir),
        "messages_scanned": len(messages),
        "downloaded_count": downloaded_total,
        "skipped_count": skipped_total,
        "failed_count": failed_total,
        "messages": message_results,
        "stream_id": stream_paths["stream_id"],
        "state_file": str(stream_paths["state_file"]),
        "ledger_file": str(stream_paths["ledger_file"]),
    }


def run_attachments_download_new(client: GraphClient, args: argparse.Namespace) -> Dict[str, Any]:
    if args.overlap_hours < 0:
        raise ValueError("--overlap-hours must be 0 or greater")
    if args.max_pages <= 0:
        raise ValueError("--max-pages must be greater than 0")
    if args.max_messages <= 0:
        raise ValueError("--max-messages must be greater than 0")

    profile = resolve_profile(args.profile)
    resolution = resolve_folder_selector(
        client=client,
        folder=args.folder,
        folder_id=args.folder_id,
        folder_path=args.folder_path,
        include_hidden=bool(args.include_hidden_folders),
    )
    account_home_id = get_authenticated_account_home_id(client.auth)
    stream_paths = build_stream_paths(
        profile=profile,
        account_home_id=account_home_id,
        folder_id=str(resolution["resolved_folder_id"] or resolution["folder_token"]),
    )
    state = load_stream_state(
        stream_paths=stream_paths,
        profile=profile,
        account_home_id=account_home_id,
        folder_id=str(resolution["resolved_folder_id"] or resolution["folder_token"]),
        folder_path=resolution.get("resolved_folder_path"),
    )
    completed_keys = load_completed_keys(stream_paths)

    run_id = uuid.uuid4().hex
    run_started = now_utc()
    run_started_iso = iso_utc(run_started)

    cursor_before = state.get("cursor_received_utc")
    cursor_before_dt = parse_graph_datetime(cursor_before)
    first_run = not bool(state.get("first_run_completed"))

    if first_run or cursor_before_dt is None:
        since_dt = run_started - timedelta(days=FIRST_RUN_BACKFILL_DAYS)
        since_source = f"first_run_backfill_{FIRST_RUN_BACKFILL_DAYS}d"
    else:
        since_dt = cursor_before_dt - timedelta(hours=int(args.overlap_hours))
        since_source = f"cursor_overlap_{int(args.overlap_hours)}h"

    since_iso = iso_utc(since_dt)

    output_dir = resolve_output_dir(args.output_dir)
    batch_dir = create_batch_output_dir(output_dir, "new")

    append_ledger_event(
        stream_paths,
        {
            "event": "run_started",
            "run_id": run_id,
            "mode": "download_new",
            "folder_id": resolution["resolved_folder_id"],
            "folder_path": resolution.get("resolved_folder_path"),
            "started_at": run_started_iso,
            "since": since_iso,
            "since_source": since_source,
            "max_pages": int(args.max_pages),
            "max_messages": int(args.max_messages),
        },
    )

    pending_before = dict(state.get("pending_failures", {}))
    retried_count = 0
    retried_success = 0
    retried_failed = 0
    retry_downloaded_total = 0
    retry_skipped_total = 0
    retry_failed_total = 0

    for key, pending in pending_before.items():
        message_id = str(pending.get("message_id") or "").strip()
        attachment_id = str(pending.get("attachment_id") or "").strip()
        if not message_id or not attachment_id:
            continue

        retried_count += 1
        try:
            metadata = client.get_attachment(message_id, attachment_id)
            retry_message = {
                "id": message_id,
                "subject": "retry_pending",
                "receivedDateTime": pending.get("last_attempt_utc"),
                "hasAttachments": True,
            }
            process = process_message_attachments(
                client=client,
                message=retry_message,
                batch_dir=batch_dir,
                completed_keys=completed_keys,
                state=state,
                stream_paths=stream_paths,
                run_id=run_id,
                force_redownload=False,
                preloaded_attachments=[metadata],
            )
            retry_downloaded_total += process["downloaded_count"]
            retry_skipped_total += process["skipped_count"]
            retry_failed_total += process["failed_count"]
            if process["failed_count"] > 0:
                retried_failed += 1
            else:
                retried_success += 1
        except (GraphAPIError, OSError, ValueError) as err:
            retried_failed += 1
            retry_failed_total += 1
            _record_pending_failure(
                state=state,
                dedupe_key=key,
                message_id=message_id,
                attachment_id=attachment_id,
                error=str(err),
                attempted_at=iso_utc(now_utc()),
            )
            append_ledger_event(
                stream_paths,
                {
                    "event": "failed",
                    "run_id": run_id,
                    "dedupe_key": key,
                    "message_id": message_id,
                    "attachment_id": attachment_id,
                    "error": str(err),
                    "phase": "pending_retry",
                },
            )

    messages = client.list_messages(
        folder=resolution["folder_token"],
        unread_only=False,
        top=int(args.max_messages),
        select_fields=list(DEFAULT_SELECT_FIELDS),
        has_attachments_only=True,
        received_since=since_iso,
        max_pages=int(args.max_pages),
    )

    message_results = []
    downloaded_total = 0
    skipped_total = 0
    failed_total = 0
    max_received_dt: Optional[datetime] = None

    for message in messages:
        result = process_message_attachments(
            client=client,
            message=message,
            batch_dir=batch_dir,
            completed_keys=completed_keys,
            state=state,
            stream_paths=stream_paths,
            run_id=run_id,
            force_redownload=False,
        )
        message_results.append(result)
        downloaded_total += result["downloaded_count"]
        skipped_total += result["skipped_count"]
        failed_total += result["failed_count"]

        received_dt = parse_graph_datetime(str(message.get("receivedDateTime") or ""))
        if received_dt and (max_received_dt is None or received_dt > max_received_dt):
            max_received_dt = received_dt

    run_completed = now_utc()
    run_completed_iso = iso_utc(run_completed)

    if max_received_dt is None:
        max_received_dt = run_started

    cursor_after_dt = max_received_dt
    if cursor_before_dt and cursor_before_dt > cursor_after_dt:
        cursor_after_dt = cursor_before_dt

    total_downloaded = retry_downloaded_total + downloaded_total
    total_skipped = retry_skipped_total + skipped_total
    total_failed = retry_failed_total + failed_total

    state["profile"] = profile
    state["account_home_id"] = account_home_id
    state["folder_id"] = str(resolution["resolved_folder_id"] or resolution["folder_token"])
    state["folder_path"] = resolution.get("resolved_folder_path")
    state["first_run_completed"] = True
    state["cursor_received_utc"] = iso_utc(cursor_after_dt)
    state["last_run_started_utc"] = run_started_iso
    state["last_run_completed_utc"] = run_completed_iso
    state["downloaded_total"] = int(state.get("downloaded_total", 0)) + total_downloaded
    state["skipped_total"] = int(state.get("skipped_total", 0)) + total_skipped
    state["failed_total"] = int(state.get("failed_total", 0)) + total_failed

    save_stream_state(stream_paths, state)

    append_ledger_event(
        stream_paths,
        {
            "event": "run_completed",
            "run_id": run_id,
            "mode": "download_new",
            "started_at": run_started_iso,
            "completed_at": run_completed_iso,
            "since": since_iso,
            "since_source": since_source,
            "messages_scanned": len(messages),
            "retried_pending": retried_count,
            "retried_success": retried_success,
            "retried_failed": retried_failed,
            "downloaded_count": total_downloaded,
            "skipped_count": total_skipped,
            "failed_count": total_failed,
            "cursor_before": cursor_before,
            "cursor_after": state["cursor_received_utc"],
        },
    )

    return {
        "run_id": run_id,
        "mode": "download_new",
        "folder_mode": resolution["folder_mode"],
        "folder_input": resolution["folder_input"],
        "resolved_folder_id": resolution["resolved_folder_id"],
        "resolved_folder_path": resolution.get("resolved_folder_path"),
        "output_dir": str(batch_dir),
        "first_run_backfill_days": FIRST_RUN_BACKFILL_DAYS if first_run else 0,
        "since": since_iso,
        "since_source": since_source,
        "cursor_before": cursor_before,
        "cursor_after": state["cursor_received_utc"],
        "max_pages": int(args.max_pages),
        "max_messages": int(args.max_messages),
        "messages_scanned": len(messages),
        "retried_pending": retried_count,
        "retried_success": retried_success,
        "retried_failed": retried_failed,
        "retry_downloaded_count": retry_downloaded_total,
        "retry_skipped_count": retry_skipped_total,
        "retry_failed_count": retry_failed_total,
        "downloaded_count": total_downloaded,
        "skipped_count": total_skipped,
        "failed_count": total_failed,
        "messages": message_results,
        "stream_id": stream_paths["stream_id"],
        "state_file": str(stream_paths["state_file"]),
        "ledger_file": str(stream_paths["ledger_file"]),
    }


def process_message_attachments(
    client: GraphClient,
    message: Dict[str, Any],
    batch_dir: Path,
    completed_keys: Set[str],
    state: Dict[str, Any],
    stream_paths: Dict[str, Any],
    run_id: str,
    force_redownload: bool,
    preloaded_attachments: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    message_id = str(message.get("id") or "").strip()
    received = message.get("receivedDateTime")
    subject = message.get("subject")

    result = {
        "message_id": message_id,
        "subject": subject,
        "received_date_time": received,
        "attachments_total": 0,
        "downloaded_count": 0,
        "skipped_count": 0,
        "failed_count": 0,
        "downloaded": [],
        "skipped": [],
        "failed": [],
    }

    if not message_id:
        result["failed_count"] = 1
        result["failed"].append({"error": "missing message id"})
        return result

    if preloaded_attachments is None:
        attachments = client.list_attachments(message_id)
    else:
        attachments = list(preloaded_attachments)

    result["attachments_total"] = len(attachments)
    if not attachments:
        return result

    message_dir = batch_dir / build_message_folder_name(message_id, subject, received)
    message_dir.mkdir(parents=True, exist_ok=True)

    for attachment in attachments:
        attachment_id = str(attachment.get("id") or "").strip()
        dedupe_key = build_dedupe_key(message_id, attachment)

        if not force_redownload and dedupe_key in completed_keys:
            result["skipped_count"] += 1
            _clear_pending_failure(state, dedupe_key)
            skipped_payload = {
                "dedupe_key": dedupe_key,
                "attachment_id": attachment_id or None,
                "name": attachment.get("name"),
                "reason": "already_downloaded",
            }
            result["skipped"].append(skipped_payload)
            append_ledger_event(
                stream_paths,
                {
                    "event": "skipped_already_downloaded",
                    "run_id": run_id,
                    "dedupe_key": dedupe_key,
                    "message_id": message_id,
                    "attachment_id": attachment_id or None,
                    "name": attachment.get("name"),
                },
            )
            continue

        if not attachment_id:
            err_msg = "attachment id missing"
            result["failed_count"] += 1
            failed_payload = {
                "dedupe_key": dedupe_key,
                "attachment_id": None,
                "name": attachment.get("name"),
                "error": err_msg,
            }
            result["failed"].append(failed_payload)
            _record_pending_failure(
                state=state,
                dedupe_key=dedupe_key,
                message_id=message_id,
                attachment_id="",
                error=err_msg,
                attempted_at=iso_utc(now_utc()),
            )
            append_ledger_event(
                stream_paths,
                {
                    "event": "failed",
                    "run_id": run_id,
                    "dedupe_key": dedupe_key,
                    "message_id": message_id,
                    "attachment_id": None,
                    "name": attachment.get("name"),
                    "error": err_msg,
                },
            )
            continue

        try:
            downloaded = download_one_attachment(
                client=client,
                message_id=message_id,
                attachment_id=attachment_id,
                target_dir=message_dir,
                metadata=attachment,
            )
            downloaded["dedupe_key"] = dedupe_key
            result["downloaded"].append(downloaded)
            result["downloaded_count"] += 1
            completed_keys.add(dedupe_key)
            _clear_pending_failure(state, dedupe_key)

            append_ledger_event(
                stream_paths,
                {
                    "event": "downloaded",
                    "run_id": run_id,
                    "dedupe_key": dedupe_key,
                    "message_id": message_id,
                    "attachment_id": attachment_id,
                    "name": attachment.get("name"),
                    "saved_path": downloaded.get("saved_path"),
                    "received_date_time": received,
                },
            )
        except (GraphAPIError, OSError, ValueError) as err:
            err_msg = str(err)
            result["failed_count"] += 1
            failed_payload = {
                "dedupe_key": dedupe_key,
                "attachment_id": attachment_id,
                "name": attachment.get("name"),
                "error": err_msg,
            }
            result["failed"].append(failed_payload)
            _record_pending_failure(
                state=state,
                dedupe_key=dedupe_key,
                message_id=message_id,
                attachment_id=attachment_id,
                error=err_msg,
                attempted_at=iso_utc(now_utc()),
            )
            append_ledger_event(
                stream_paths,
                {
                    "event": "failed",
                    "run_id": run_id,
                    "dedupe_key": dedupe_key,
                    "message_id": message_id,
                    "attachment_id": attachment_id,
                    "name": attachment.get("name"),
                    "error": err_msg,
                },
            )

    return result


def build_auth_manager(profile: Optional[str]) -> AuthManager:
    config = AuthConfig.from_env(profile_override=profile)
    return AuthManager(config)


def build_graph_client(profile: Optional[str]) -> GraphClient:
    manager = build_auth_manager(profile)
    return GraphClient(manager)


def build_onboarding_plan(args: argparse.Namespace) -> Dict[str, Any]:
    profile = (args.profile or os.environ.get("OUTLOOK_PROFILE") or "default").strip() or "default"
    client_id = (args.client_id or os.environ.get("OUTLOOK_CLIENT_ID") or "").strip()
    tenant_id = (args.tenant_id or os.environ.get("OUTLOOK_TENANT_ID") or "common").strip() or "common"
    redirect_uri = (
        args.redirect_uri
        or os.environ.get("OUTLOOK_REDIRECT_URI")
        or "http://localhost:8765"
    ).strip()
    scopes = normalize_scope_list(args.scopes or os.environ.get("OUTLOOK_SCOPES"))

    questions_for_user: List[str] = []
    if not client_id:
        questions_for_user.append(
            "Please share your Microsoft Entra Application (client) ID for the Outlook app registration."
        )
    if not args.tenant_id and not os.environ.get("OUTLOOK_TENANT_ID"):
        questions_for_user.append(
            "Should we use tenant mode 'common' (multi-tenant/personal) or a specific tenant ID?"
        )

    user_actions = [
        "In Entra App Registration > Authentication, add platform 'Mobile and desktop applications'.",
        "Set redirect URI to http://localhost:8765.",
        "Enable 'Allow public client flows'.",
        "In API permissions, grant delegated permissions: User.Read, Mail.ReadWrite, Mail.Send.",
        "Run login command and complete browser sign-in/consent once.",
    ]
    if tenant_id == "common":
        user_actions.insert(
            0,
            "Ensure app account type supports multi-tenant/personal access when using tenant_id=common.",
        )

    cli_script = shlex.quote(str(Path(__file__).resolve()))
    env_prefix = (
        f'OUTLOOK_CLIENT_ID="{client_id or "<CLIENT_ID>"}" '
        f'OUTLOOK_TENANT_ID="{tenant_id}" '
        f'OUTLOOK_REDIRECT_URI="{redirect_uri}" '
        f'OUTLOOK_SCOPES="{" ".join(scopes)}" '
    )
    login_cmd = (
        f"{env_prefix}"
        f"python3 {cli_script} auth login --method {args.method} --profile {profile}"
    )
    status_cmd = (
        f"{env_prefix}"
        f"python3 {cli_script} auth status --profile {profile}"
    )
    first_mail_cmd = (
        f"{env_prefix}"
        f"python3 {cli_script} mail list --folder inbox --unread-only --top 10 --profile {profile}"
    )

    status: Optional[Dict[str, Any]] = None
    authenticated = False
    if client_id:
        config = AuthConfig(
            client_id=client_id,
            tenant_id=tenant_id,
            redirect_uri=redirect_uri,
            scopes=scopes,
            profile=profile,
            token_store_mode=os.environ.get("OUTLOOK_TOKEN_STORE", "auto").strip().lower() or "auto",
        )
        status = AuthManager(config).status()
        authenticated = bool(status.get("authenticated"))

    return {
        "profile": profile,
        "config": {
            "client_id_configured": bool(client_id),
            "tenant_id": tenant_id,
            "redirect_uri": redirect_uri,
            "scopes": scopes,
        },
        "questions_for_user": questions_for_user,
        "required_user_actions": user_actions,
        "agent_next_steps": [
            "Collect any missing fields from questions_for_user.",
            "Share required_user_actions with the user in concise bullets.",
            "Run login_command and wait for user to complete browser/device consent.",
            "Run status_command and confirm authenticated=true before any mail operations.",
        ],
        "login_command": login_cmd,
        "status_command": status_cmd,
        "first_mail_command": first_mail_cmd,
        "ready_for_login": bool(client_id),
        "already_authenticated": authenticated,
        "status": status,
    }


def resolve_profile(profile_override: Optional[str]) -> str:
    return (profile_override or os.environ.get("OUTLOOK_PROFILE") or "default").strip() or "default"


def get_authenticated_account_home_id(manager: AuthManager) -> str:
    status = manager.status()
    account = status.get("account") if isinstance(status, dict) else None
    home_account_id = account.get("home_account_id") if isinstance(account, dict) else None
    if not home_account_id:
        raise AuthError("No authenticated account found for this profile. Run auth login first.")
    return str(home_account_id)


def parse_select_fields(raw: str) -> List[str]:
    fields = [field.strip() for field in raw.split(",") if field.strip()]
    return fields or list(DEFAULT_SELECT_FIELDS)


def normalize_scope_list(raw: Optional[str]) -> List[str]:
    source = raw or " ".join(DEFAULT_AUTH_SCOPES)
    scopes: List[str] = []
    seen = set()
    for chunk in source.replace(",", " ").split():
        scope = chunk.strip()
        if not scope:
            continue
        if scope.lower() in RESERVED_OIDC_SCOPES:
            continue
        lowered = scope.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        scopes.append(scope)
    return scopes or list(DEFAULT_AUTH_SCOPES)


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


def resolve_folder_selector(
    client: GraphClient,
    folder: Optional[str],
    folder_id: Optional[str],
    folder_path: Optional[str],
    include_hidden: bool,
) -> Dict[str, Any]:
    if folder_id:
        raw = str(folder_id).strip()
        if not raw:
            raise ValueError("--folder-id cannot be empty")
        resolved_id = raw
        resolved_path = None
        try:
            meta = client.get_mail_folder(raw)
            resolved_id = str(meta.get("id") or raw)
            display_name = str(meta.get("displayName") or "").strip()
            if display_name:
                resolved_path = f"/{display_name}"
        except GraphAPIError:
            pass

        return {
            "folder_mode": "id",
            "folder_input": raw,
            "folder_token": resolved_id,
            "resolved_folder_id": resolved_id,
            "resolved_folder_path": resolved_path,
        }

    if folder_path:
        return resolve_folder_path_selector(client, folder_path, include_hidden)

    token = (folder or DEFAULT_FOLDER_ROOT).strip() or DEFAULT_FOLDER_ROOT
    resolved_id = token
    resolved_path = None
    try:
        meta = client.get_mail_folder(token)
        resolved_id = str(meta.get("id") or token)
        display_name = str(meta.get("displayName") or "").strip()
        if display_name:
            resolved_path = f"/{display_name}"
    except GraphAPIError:
        pass

    return {
        "folder_mode": "token",
        "folder_input": token,
        "folder_token": token,
        "resolved_folder_id": resolved_id,
        "resolved_folder_path": resolved_path,
    }


def resolve_folder_path_selector(client: GraphClient, raw_path: str, include_hidden: bool) -> Dict[str, Any]:
    root_meta = client.get_mail_folder("inbox")
    root_id = str(root_meta.get("id") or "").strip()
    if not root_id:
        raise ValueError("Could not resolve Inbox root folder id")

    root_name = str(root_meta.get("displayName") or "Inbox").strip() or "Inbox"
    current_id = root_id
    current_path = f"/{root_name}"

    segments = normalize_folder_path(raw_path)
    for segment in segments:
        children = client.list_child_folders(current_id, include_hidden=include_hidden)
        matches = [
            child
            for child in children
            if str(child.get("displayName") or "").strip().casefold() == segment.casefold()
        ]

        if not matches:
            available = sorted(
                {
                    str(child.get("displayName") or "").strip()
                    for child in children
                    if str(child.get("displayName") or "").strip()
                }
            )
            preview = ", ".join(available[:20]) if available else "<none>"
            raise ValueError(
                f"Folder path segment '{segment}' not found under '{current_path}'. "
                f"Available children: {preview}"
            )

        if len(matches) > 1:
            ids = [str(child.get("id") or "") for child in matches]
            raise ValueError(
                f"Folder path segment '{segment}' is ambiguous under '{current_path}'. "
                f"Matching folder ids: {ids}"
            )

        selected = matches[0]
        current_id = str(selected.get("id") or "").strip()
        if not current_id:
            raise ValueError(f"Resolved folder segment '{segment}' has no id")
        display_name = str(selected.get("displayName") or segment)
        current_path = f"{current_path}/{display_name}"

    return {
        "folder_mode": "path",
        "folder_input": raw_path,
        "folder_token": current_id,
        "resolved_folder_id": current_id,
        "resolved_folder_path": current_path,
    }


def normalize_folder_path(raw_path: str) -> List[str]:
    normalized = (raw_path or "").strip()
    if not normalized:
        return []

    normalized = normalized.strip("/")
    if not normalized:
        return []

    pieces = [piece.strip() for piece in normalized.split("/") if piece.strip()]
    if pieces and pieces[0].casefold() == "inbox":
        pieces = pieces[1:]
    return pieces


def build_folder_tree(
    client: GraphClient,
    root_token: str,
    include_hidden: bool,
    max_nodes: int,
) -> Dict[str, Any]:
    root_meta = client.get_mail_folder(root_token)
    root_name = str(root_meta.get("displayName") or "Inbox").strip() or "Inbox"
    root_path = f"/{root_name}"

    index_rows: List[Dict[str, Any]] = []
    node_counter = 0

    def walk(meta: Dict[str, Any], path: str) -> Dict[str, Any]:
        nonlocal node_counter
        node_counter += 1
        if node_counter > max_nodes:
            raise ValueError(
                f"Folder traversal exceeded max nodes ({max_nodes}). "
                "Rerun with a higher --max-nodes value."
            )

        node = to_folder_node(meta, path)
        index_rows.append({k: node[k] for k in DEFAULT_FOLDER_FIELDS + ["path"]})

        child_count = int(meta.get("childFolderCount") or 0)
        if child_count <= 0:
            return node

        children = client.list_child_folders(str(meta.get("id") or ""), include_hidden=include_hidden)
        ordered = sorted(
            children,
            key=lambda item: (
                str(item.get("displayName") or "").casefold(),
                str(item.get("id") or ""),
            ),
        )

        for child in ordered:
            name = str(child.get("displayName") or child.get("id") or "unknown")
            child_path = f"{path}/{name}"
            node["children"].append(walk(child, child_path))

        return node

    tree = walk(root_meta, root_path)
    return {"tree": tree, "index": index_rows}


def to_folder_node(meta: Dict[str, Any], path: str) -> Dict[str, Any]:
    return {
        "id": str(meta.get("id") or ""),
        "display_name": meta.get("displayName"),
        "parent_id": meta.get("parentFolderId"),
        "child_folder_count": int(meta.get("childFolderCount") or 0),
        "total_item_count": int(meta.get("totalItemCount") or 0),
        "unread_item_count": int(meta.get("unreadItemCount") or 0),
        "is_hidden": bool(meta.get("isHidden", False)),
        "path": path,
        "children": [],
    }


def get_state_base_dir() -> Path:
    override = os.environ.get("OUTLOOK_STATE_DIR")
    if override:
        base = Path(override).expanduser().resolve()
    else:
        base = Path(__file__).resolve().parent.parent / ".state"
    base.mkdir(parents=True, exist_ok=True)
    return base


def build_stream_paths(profile: str, account_home_id: str, folder_id: str) -> Dict[str, Any]:
    stream_seed = f"{profile}|{account_home_id}|{folder_id}"
    stream_id = hashlib.sha256(stream_seed.encode("utf-8")).hexdigest()[:24]
    stream_dir = get_state_base_dir() / stream_id
    return {
        "stream_id": stream_id,
        "stream_dir": stream_dir,
        "state_file": stream_dir / "state.json",
        "ledger_file": stream_dir / "ledger.jsonl",
    }


def default_stream_state(
    profile: str,
    account_home_id: str,
    folder_id: str,
    folder_path: Optional[str],
) -> Dict[str, Any]:
    return {
        "version": STATE_VERSION,
        "profile": profile,
        "account_home_id": account_home_id,
        "folder_id": folder_id,
        "folder_path": folder_path,
        "first_run_completed": False,
        "cursor_received_utc": None,
        "last_run_started_utc": None,
        "last_run_completed_utc": None,
        "pending_failures": {},
        "downloaded_total": 0,
        "skipped_total": 0,
        "failed_total": 0,
    }


def load_stream_state(
    stream_paths: Dict[str, Any],
    profile: str,
    account_home_id: str,
    folder_id: str,
    folder_path: Optional[str],
) -> Dict[str, Any]:
    state = default_stream_state(profile, account_home_id, folder_id, folder_path)
    state_file: Path = stream_paths["state_file"]

    if state_file.exists():
        try:
            loaded = json.loads(state_file.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                state.update(loaded)
        except json.JSONDecodeError as err:
            raise ValueError(f"Invalid JSON in state file '{state_file}': {err}") from err

    if not isinstance(state.get("pending_failures"), dict):
        state["pending_failures"] = {}

    state["version"] = STATE_VERSION
    state["profile"] = profile
    state["account_home_id"] = account_home_id
    state["folder_id"] = folder_id
    if folder_path:
        state["folder_path"] = folder_path

    return state


def save_stream_state(stream_paths: Dict[str, Any], state: Dict[str, Any]) -> None:
    stream_dir: Path = stream_paths["stream_dir"]
    state_file: Path = stream_paths["state_file"]
    stream_dir.mkdir(parents=True, exist_ok=True)

    tmp_file = state_file.with_suffix(".tmp")
    tmp_file.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")
    tmp_file.replace(state_file)


def append_ledger_event(stream_paths: Dict[str, Any], payload: Dict[str, Any]) -> None:
    stream_dir: Path = stream_paths["stream_dir"]
    ledger_file: Path = stream_paths["ledger_file"]
    stream_dir.mkdir(parents=True, exist_ok=True)

    event = dict(payload)
    event.setdefault("ts", iso_utc(now_utc()))

    with ledger_file.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(event, sort_keys=True))
        handle.write("\n")


def load_completed_keys(stream_paths: Dict[str, Any]) -> Set[str]:
    ledger_file: Path = stream_paths["ledger_file"]
    completed: Set[str] = set()
    if not ledger_file.exists():
        return completed

    with ledger_file.open("r", encoding="utf-8") as handle:
        for line in handle:
            raw = line.strip()
            if not raw:
                continue
            try:
                event = json.loads(raw)
            except json.JSONDecodeError:
                continue
            if event.get("event") != "downloaded":
                continue
            key = str(event.get("dedupe_key") or "").strip()
            if key:
                completed.add(key)

    return completed


def count_ledger_entries(stream_paths: Dict[str, Any]) -> int:
    ledger_file: Path = stream_paths["ledger_file"]
    if not ledger_file.exists():
        return 0
    count = 0
    with ledger_file.open("r", encoding="utf-8") as handle:
        for _ in handle:
            count += 1
    return count


def _record_pending_failure(
    state: Dict[str, Any],
    dedupe_key: str,
    message_id: str,
    attachment_id: str,
    error: str,
    attempted_at: str,
) -> None:
    pending = state.setdefault("pending_failures", {})
    if not isinstance(pending, dict):
        pending = {}
        state["pending_failures"] = pending

    previous = pending.get(dedupe_key, {}) if isinstance(pending.get(dedupe_key), dict) else {}
    attempts = int(previous.get("attempts", 0)) + 1

    pending[dedupe_key] = {
        "message_id": message_id,
        "attachment_id": attachment_id,
        "attempts": attempts,
        "last_error": error,
        "last_attempt_utc": attempted_at,
    }


def _clear_pending_failure(state: Dict[str, Any], dedupe_key: str) -> None:
    pending = state.get("pending_failures")
    if isinstance(pending, dict):
        pending.pop(dedupe_key, None)


def build_dedupe_key(message_id: str, attachment: Dict[str, Any]) -> str:
    attachment_id = str(attachment.get("id") or "").strip()
    if attachment_id:
        return f"{message_id}:{attachment_id}"

    name = str(attachment.get("name") or "attachment")
    size = attachment.get("size")
    return f"{message_id}:{name}:{size}"


def now_utc() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def iso_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_graph_datetime(raw: Optional[str]) -> Optional[datetime]:
    if not raw:
        return None

    normalized = raw.strip()
    if not normalized:
        return None

    if normalized.endswith("Z"):
        normalized = normalized[:-1] + "+00:00"

    try:
        return datetime.fromisoformat(normalized).astimezone(timezone.utc)
    except ValueError:
        return None


def create_batch_output_dir(parent: Path, mode: str) -> Path:
    timestamp = now_utc().strftime("%Y%m%dT%H%M%SZ")
    folder = parent / f"{mode}_{timestamp}"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def build_message_folder_name(message_id: str, subject: Any, received: Any) -> str:
    received_dt = parse_graph_datetime(str(received or ""))
    received_token = received_dt.strftime("%Y%m%dT%H%M%SZ") if received_dt else "unknown_time"

    subject_text = sanitize_filename(str(subject or "no_subject"))
    subject_text = subject_text[:80].strip(" _") or "no_subject"

    id_token = sanitize_filename(message_id)[:16] or "message"
    return f"msg_{received_token}_{subject_text}_{id_token}"


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
