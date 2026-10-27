"""Command line: run the server, sign in/out, and set up Claude Desktop."""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import sys
from pathlib import Path

from .auth import AuthError, Authenticator, SignInRequired
from .config import ConfigError, app_dir, load_settings

SERVER_NAME = "onenote"


def _setup_logging() -> None:
    # stdout carries MCP JSON-RPC, so all logging goes to stderr.
    logging.basicConfig(
        stream=sys.stderr,
        level=os.environ.get("ONENOTE_LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    for noisy in ("httpx", "httpcore", "msal"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def cmd_serve(_: argparse.Namespace) -> int:
    from .server import build_server

    _setup_logging()
    build_server().run(transport="stdio")
    return 0


def _auth() -> Authenticator:
    return Authenticator(load_settings())


def cmd_login(_: argparse.Namespace) -> int:
    auth = _auth()
    try:
        auth.get_token()
        print(f"Already signed in as {auth.signed_in_user()}.")
        return 0
    except SignInRequired:
        pass
    user = auth.login_interactive(lambda msg: print(msg, flush=True))
    print(f"Signed in as {user}. Token cache: {auth.settings.token_cache_path}")
    return 0


def cmd_status(_: argparse.Namespace) -> int:
    settings = load_settings()
    print(f"Client ID:   {settings.client_id}")
    print(f"Authority:   {settings.authority_url}")
    print(f"Token cache: {settings.token_cache_path}")
    mode = "read-only"
    if settings.flags.allow_write:
        mode = "read/write" + (" + delete" if settings.flags.allow_delete else "")
    print(f"Mode:        {mode} (scope {settings.scopes[0].rsplit('/', 1)[-1]})")
    auth = Authenticator(settings)
    try:
        auth.get_token()
        print(f"Signed in:   yes ({auth.signed_in_user()})")
    except SignInRequired:
        print("Signed in:   no (run `onenote-mcp login`)")
    return 0


def cmd_logout(_: argparse.Namespace) -> int:
    auth = _auth()
    auth.logout()
    print(f"Signed out; removed {auth.settings.token_cache_path}.")
    return 0


def desktop_config_path() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
    if sys.platform == "win32":
        return Path(os.environ["APPDATA"]) / "Claude" / "claude_desktop_config.json"
    return Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "Claude" / "claude_desktop_config.json"


def server_entry() -> dict:
    """Launch with this interpreter so Claude Desktop needn't find uv on its PATH."""
    settings = load_settings()
    env = {"ONENOTE_CLIENT_ID": settings.client_id, "ONENOTE_AUTHORITY": settings.authority}
    for key in ("ONENOTE_TOKEN_CACHE", "ONENOTE_MCP_HOME", "ONENOTE_ALLOW_WRITE", "ONENOTE_ALLOW_DELETE"):
        if os.environ.get(key):
            env[key] = os.environ[key]
    return {"command": sys.executable, "args": ["-m", "onenote_mcp"], "env": env}


def cmd_desktop_config(args: argparse.Namespace) -> int:
    entry = server_entry()
    if not args.write:
        print(json.dumps({"mcpServers": {args.name: entry}}, indent=2))
        print(f"\nMerge this into {desktop_config_path()} or rerun with --write.", file=sys.stderr)
        return 0

    path = Path(args.path) if args.path else desktop_config_path()
    config: dict = {}
    if path.exists():
        try:
            config = json.loads(path.read_text(encoding="utf-8") or "{}")
        except json.JSONDecodeError as exc:
            print(f"{path} is not valid JSON ({exc}); not changing it.", file=sys.stderr)
            return 1
        backup = path.with_name(path.name + ".bak")
        shutil.copy2(path, backup)
        print(f"Backed up existing config to {backup}")
    config.setdefault("mcpServers", {})[args.name] = entry
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(config, indent=2) + "\n", encoding="utf-8")
    print(f"Added '{args.name}' to {path}. Restart Claude Desktop to load it.")
    return 0


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(prog="onenote-mcp", description="Read-only OneNote MCP server")
    sub = parser.add_subparsers(dest="command")
    sub.add_parser("serve", help="Run the MCP server over stdio (default)").set_defaults(fn=cmd_serve)
    sub.add_parser("login", help="Sign in with the device code flow").set_defaults(fn=cmd_login)
    sub.add_parser("status", help="Show configuration and sign-in state").set_defaults(fn=cmd_status)
    sub.add_parser("logout", help="Forget cached tokens").set_defaults(fn=cmd_logout)
    dc = sub.add_parser("desktop-config", help="Print or write the Claude Desktop config entry")
    dc.add_argument("--write", action="store_true", help="Merge into claude_desktop_config.json")
    dc.add_argument("--path", help="Config file to write (defaults to Claude Desktop's)")
    dc.add_argument("--name", default=SERVER_NAME, help="Server name in the config")
    dc.set_defaults(fn=cmd_desktop_config)

    args = parser.parse_args(argv)
    fn = getattr(args, "fn", cmd_serve)
    try:
        sys.exit(fn(args))
    except (ConfigError, AuthError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
