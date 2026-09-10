#!/usr/bin/env python3
"""
migrate_archive.py

Uploads Thunderbird mbox folder trees (the classic "Foldername" + "Foldername.sbd"
pattern) straight into an Exchange Online mailbox over IMAP, recreating the folder
hierarchy as it goes. Written to work around Thunderbird's own drag-and-drop copy,
which was failing (TRYCREATE errors) on deeply nested folders.

Supports migrating multiple source accounts/profiles in a single run, each landing
under its own top-level folder in the destination mailbox - see CONFIGURATION below.

Requires only the Python standard library - nothing to install.

HOW IT AUTHENTICATES
--------------------
Uses OAuth2 "device code" flow against your own tenant's app registration.
When you run this, it will print a URL and a short code. Open the URL in any
browser, sign in as the mailbox owner, and enter the code. Your password
never touches this script - only Microsoft's own login page sees it, and the
script only ever receives a short-lived access token back.

Access tokens only last 60-90 minutes, and Exchange Online closes the IMAP
connection once yours expires. On a run this long that WILL happen at least
once. The script detects the dropped connection, silently uses the refresh
token from your original sign-in to get a new access token (no re-scanning a
code), reconnects, and carries on with the same folder it was in the middle
of - no need to babysit it.

HOW IT TRACKS PROGRESS
-----------------------
Every folder it fully finishes gets recorded in migration_state.json (created next
to this script). Re-running the script skips folders already marked done, so it's
safe to stop (Ctrl+C) and resume later, or to just re-run after a crash. Progress is
tracked per destination folder path, so this works the same whether you have one
source or several.

CONFIGURATION
-------------
This script reads its settings (tenant ID, app registration client ID, mailbox
username, the source folder(s) to migrate, and any folder names to skip) from a
config.json file rather than hardcoding them, since those are specific to your own
tenant/mailbox and shouldn't be committed to a public repo.

    1. Copy config.example.json to ../private/config.json (relative to this
       script) - the "private" folder is git-ignored, so anything you put
       there never gets committed.
    2. Fill in your own tenant_id, client_id, and username.
    3. Fill in "sources": a list of {"source_dir": ..., "dest_top_folder": ...}
       pairs. Each source_dir is a directory holding one or more Thunderbird
       top-level mbox files directly (for example a Thunderbird "Mail/<host>"
       account folder, or a "Local Folders" directory) - every top-level mbox
       file found directly inside it (INBOX, Sent, Archives, etc., each with
       its own optional "<name>.sbd" subfolder tree) is recreated, with its
       full subfolder structure, under dest_top_folder in the destination
       mailbox. List as many sources as you have accounts to migrate; they
       all land in the same mailbox (USERNAME), each under its own
       dest_top_folder, and all share one migration_state.json.
    4. Optionally set "skip_folder_names": a list of folder names (matched
       case-insensitively, at any depth in any source) to leave out of the
       migration entirely - the folder itself and everything nested under it
       (e.g. ["Trash"] to skip every Trash folder and its contents).
    5. Run the script as usual - it will pick up ../private/config.json
       automatically if present, falling back to config.example.json
       (with its placeholder values) otherwise.

USAGE
-----
    python migrate_archive.py                  # full run, all configured sources
    python migrate_archive.py --dry-run         # list what would happen, no changes
    python migrate_archive.py --only "2010"     # only migrate folders whose destination path contains "2010"
    python migrate_archive.py --reset           # ignore the state file and start clean
"""

import argparse
import base64
import email.utils
import imaplib
import json
import mailbox
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

_LIST_RE = re.compile(rb'^\((?P<flags>[^)]*)\)\s+"(?P<delim>[^"]*)"\s+(?P<name>.+)$')


def _parse_list_line(entry):
    """Parse one line of an IMAP LIST response into (delimiter, mailbox_name)."""
    m = _LIST_RE.match(entry)
    if not m:
        return None, None
    delim = m.group("delim").decode(errors="replace")
    name_raw = m.group("name").decode(errors="replace").strip()
    if name_raw.startswith('"') and name_raw.endswith('"'):
        name_raw = name_raw[1:-1]
    return delim, name_raw

# ============================== CONFIG ======================================
# Real values live in ../private/config.json (git-ignored, not in this repo).
# See config.example.json for the schema and the "CONFIGURATION" note above.

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PRIVATE_CONFIG = os.path.join(_SCRIPT_DIR, "..", "private", "config.json")
_EXAMPLE_CONFIG = os.path.join(_SCRIPT_DIR, "config.example.json")


def _load_config():
    path = _PRIVATE_CONFIG if os.path.exists(_PRIVATE_CONFIG) else _EXAMPLE_CONFIG
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if path == _EXAMPLE_CONFIG:
        print(
            "NOTE: no ../private/config.json found - using config.example.json's "
            "placeholder values. Copy it to ../private/config.json and fill in your "
            "own tenant/mailbox details before running for real.\n"
        )
    return cfg


_cfg = _load_config()

TENANT_ID = _cfg["tenant_id"]           # your Azure AD tenant ID
CLIENT_ID = _cfg["client_id"]           # your "device code" app registration's client ID
USERNAME = _cfg["username"]             # mailbox to sign in as / migrate into

# List of {"source_dir": ..., "dest_top_folder": ...} pairs. Each source_dir is a
# directory holding one or more Thunderbird top-level mbox files directly (plus
# their optional "<name>.sbd" subfolder trees); everything found in it is recreated
# under dest_top_folder in the destination mailbox. One entry per account/profile
# being migrated - see the CONFIGURATION note above.
SOURCES = _cfg["sources"]

# Folder names (case-insensitive, matched at any depth in any source) to leave out
# of the migration entirely - the folder itself and everything nested under it.
SKIP_FOLDER_NAMES = {name.lower() for name in _cfg.get("skip_folder_names", [])}

IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993
SCOPE = "https://outlook.office365.com/IMAP.AccessAsUser.All offline_access"

STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "migration_state.json")
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "migration_log.txt")

# Names Thunderbird creates alongside a real mbox folder that are NOT mail data -
# safe to ignore entirely.
SKIP_SUFFIXES = (".msf", ".mozmsgs")

APPEND_RETRY_ATTEMPTS = 5
APPEND_RETRY_BASE_DELAY = 2.0   # seconds, doubles each retry (exponential backoff)
APPEND_PACING_DELAY = 0.15      # small delay between appends to be gentle on the server

MAX_RECONNECTS = 3              # how many times to reconnect+retry a single IMAP op
RECONNECT_PAUSE = 3.0           # seconds to wait before reconnecting

# ============================================================================


def log(msg):
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ----------------------------- OAuth2 device code flow -----------------------

def get_access_token():
    """Full device-code sign-in. Returns (access_token, refresh_token) - the
    refresh_token (present because SCOPE includes offline_access) lets later
    reconnects renew the access token silently, without prompting again."""
    devicecode_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/devicecode"
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    data = urllib.parse.urlencode({"client_id": CLIENT_ID, "scope": SCOPE}).encode()
    req = urllib.request.Request(devicecode_url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        payload = json.loads(resp.read().decode())

    print("\n" + "=" * 70)
    print(payload["message"])
    print("=" * 70 + "\n")

    interval = payload.get("interval", 5)
    device_code = payload["device_code"]
    expires_in = payload.get("expires_in", 900)
    deadline = time.time() + expires_in

    while time.time() < deadline:
        time.sleep(interval)
        data = urllib.parse.urlencode({
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": CLIENT_ID,
            "device_code": device_code,
        }).encode()
        req = urllib.request.Request(token_url, data=data, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                tok = json.loads(resp.read().decode())
                log("Signed in successfully.")
                return tok["access_token"], tok.get("refresh_token")
        except urllib.error.HTTPError as e:
            body = json.loads(e.read().decode())
            err = body.get("error")
            if err == "authorization_pending":
                continue
            elif err == "slow_down":
                interval += 5
                continue
            else:
                raise RuntimeError(f"Device code sign-in failed: {body}")

    raise RuntimeError("Device code expired before you signed in - run the script again.")


def refresh_access_token(refresh_token):
    """Silently exchange a refresh token for a new access token - no browser,
    no code. Returns (access_token, refresh_token); AAD sometimes rotates the
    refresh token, so always keep whichever one comes back (falling back to
    the one we sent if none is returned)."""
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "refresh_token": refresh_token,
        "scope": SCOPE,
    }).encode()
    req = urllib.request.Request(token_url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        tok = json.loads(resp.read().decode())
    return tok["access_token"], tok.get("refresh_token", refresh_token)


# ----------------------------- IMAP helpers -----------------------------------

def imap_connect(access_token):
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    auth_string = f"user={USERNAME}\x01auth=Bearer {access_token}\x01\x01"
    imap.authenticate("XOAUTH2", lambda _: auth_string.encode())
    return imap


class ImapSession:
    """Holds the live IMAP connection plus the OAuth refresh token, and knows
    how to repair itself when the connection drops - which happens routinely
    on a run this long, typically right around when the ~60-90 minute access
    token expires and Exchange Online closes the socket on us."""

    def __init__(self):
        self.refresh_token = None
        self.imap = None
        self._connect()

    def _connect(self):
        if self.refresh_token:
            try:
                access_token, self.refresh_token = refresh_access_token(self.refresh_token)
            except Exception as e:
                log(f"Silent token refresh failed ({e}); falling back to a fresh sign-in.")
                log("Requesting device code sign-in...")
                access_token, self.refresh_token = get_access_token()
        else:
            log("Requesting device code sign-in...")
            access_token, self.refresh_token = get_access_token()
        self.imap = imap_connect(access_token)

    def reconnect(self):
        try:
            self.imap.logout()
        except Exception:
            pass
        time.sleep(RECONNECT_PAUSE)
        self._connect()
        log("Reconnected.")


def call_with_reconnect(session, fn, *args, **kwargs):
    """Call fn(session.imap, *args, **kwargs). If the connection has dropped
    (imaplib abort / a raw socket error), reconnect - refreshing the access
    token silently if possible - and retry, up to MAX_RECONNECTS times."""
    last_err = None
    for attempt in range(MAX_RECONNECTS + 1):
        try:
            return fn(session.imap, *args, **kwargs)
        except (imaplib.IMAP4.abort, ConnectionError, OSError) as e:
            last_err = e
            if attempt == MAX_RECONNECTS:
                break
            log(f"Connection dropped ({e}); reconnecting (attempt {attempt + 1}/{MAX_RECONNECTS})...")
            session.reconnect()
    raise RuntimeError(f"Giving up after {MAX_RECONNECTS} reconnect attempts: {last_err}")


def get_delimiter(imap):
    typ, data = imap.list('""', '"INBOX"')
    if typ == "OK" and data and data[0]:
        delim, _name = _parse_list_line(data[0])
        if delim:
            return delim
    return "/"


_existing_folders_cache = None


def folder_exists(imap, path):
    global _existing_folders_cache
    if _existing_folders_cache is None:
        _existing_folders_cache = set()
        typ, data = imap.list()
        if typ == "OK":
            for entry in data:
                if not entry:
                    continue
                _delim, name = _parse_list_line(entry)
                if name:
                    _existing_folders_cache.add(name)
    return path in _existing_folders_cache


def ensure_folder(imap, path, dry_run):
    if folder_exists(imap, path):
        return
    if dry_run:
        log(f"[dry-run] would CREATE folder: {path}")
        _existing_folders_cache.add(path)
        return
    typ, data = imap.create(f'"{path}"')
    if typ != "OK":
        raise RuntimeError(f"Could not create folder {path}: {data}")
    _existing_folders_cache.add(path)
    log(f"Created folder: {path}")


def append_message(imap, folder_path, raw_bytes, internaldate):
    last_err = None
    for attempt in range(1, APPEND_RETRY_ATTEMPTS + 1):
        try:
            typ, data = imap.append(f'"{folder_path}"', None, internaldate, raw_bytes)
            if typ == "OK":
                return
            last_err = data
        except imaplib.IMAP4.abort as e:
            # connection dropped - bubble up so caller can reconnect
            raise
        except Exception as e:
            last_err = e
        delay = APPEND_RETRY_BASE_DELAY * (2 ** (attempt - 1))
        log(f"  append retry {attempt}/{APPEND_RETRY_ATTEMPTS} after error: {last_err} (sleeping {delay:.0f}s)")
        time.sleep(delay)
    raise RuntimeError(f"Giving up on a message in {folder_path}: {last_err}")


# ----------------------------- state tracking ----------------------------------

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"folders_done": {}, "folder_progress": {}}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


# ----------------------------- mbox walking ------------------------------------

def internaldate_for(raw_bytes):
    try:
        import email
        msg = email.message_from_bytes(raw_bytes)
        date_hdr = msg.get("Date")
        if date_hdr:
            dt = email.utils.parsedate_to_datetime(date_hdr)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=email.utils.datetime.timezone.utc)
            return imaplib.Time2Internaldate(dt.timestamp())
    except Exception:
        pass
    return imaplib.Time2Internaldate(time.time())


def migrate_one_folder(session, local_folder_file, imap_path, state, dry_run):
    """Append all messages in a single mbox file to imap_path, resuming if needed."""
    if state["folders_done"].get(imap_path):
        log(f"Skipping (already done): {imap_path}")
        return

    # Always create the folder itself, even if it holds no direct messages -
    # its children (if any) need a real parent to nest under.
    call_with_reconnect(session, ensure_folder, imap_path, dry_run)

    if not os.path.exists(local_folder_file) or os.path.getsize(local_folder_file) == 0:
        state["folders_done"][imap_path] = True
        save_state(state)
        return

    start_index = state["folder_progress"].get(imap_path, 0)
    # Opened for reading only - we never call any mutating method (add/remove/
    # flush) on this box, so the original file on disk is never rewritten.
    # We deliberately never call box.close() either, since mailbox.mbox.close()
    # would flush; leaving the handle open for the process lifetime is harmless
    # for a script that only ever reads a few hundred files sequentially.
    box = mailbox.mbox(local_folder_file, factory=None, create=False)
    keys = list(box.keys())
    total = len(keys)
    log(f"{imap_path}: {total} message(s) in source, resuming at {start_index}")
    for i, key in enumerate(keys):
        if i < start_index:
            continue
        raw = box.get_bytes(key)
        idate = internaldate_for(raw)
        if dry_run:
            log(f"[dry-run] would APPEND message {i + 1}/{total} to {imap_path}")
        else:
            call_with_reconnect(session, append_message, imap_path, raw, idate)
            time.sleep(APPEND_PACING_DELAY)
        state["folder_progress"][imap_path] = i + 1
        if (i + 1) % 25 == 0:
            save_state(state)
            log(f"{imap_path}: {i + 1}/{total}")

    state["folders_done"][imap_path] = True
    save_state(state)
    log(f"Finished: {imap_path} ({total} messages)")


def _is_skipped(name):
    return name.lower() in SKIP_FOLDER_NAMES


def walk_and_migrate(session, local_dir, local_name, imap_parent_path, delimiter, state, dry_run, only_filter):
    """
    local_dir: directory containing `local_name` (mbox file) and `local_name.sbd` (subfolder dir), if any.
    """
    imap_path = f"{imap_parent_path}{delimiter}{local_name}" if imap_parent_path else local_name

    mbox_file = os.path.join(local_dir, local_name)
    sbd_dir = os.path.join(local_dir, local_name + ".sbd")

    do_this_folder = (not only_filter) or (only_filter in imap_path)

    if do_this_folder:
        migrate_one_folder(session, mbox_file, imap_path, state, dry_run)

    if os.path.isdir(sbd_dir):
        for entry in sorted(os.listdir(sbd_dir)):
            if entry.endswith(SKIP_SUFFIXES):
                continue
            full = os.path.join(sbd_dir, entry)
            if entry.endswith(".sbd"):
                continue  # handled as the sibling of its base name
            if os.path.isdir(full):
                # a directory that isn't ".sbd" and isn't a skip suffix - shouldn't
                # normally happen in a Thunderbird tree, but don't choke on it
                continue
            if _is_skipped(entry):
                # Skip the folder AND everything nested under it - just don't
                # recurse into it at all, so its own .sbd subtree (if any) is
                # never even looked at.
                log(f"Skipping folder (configured skip list): {imap_path}{delimiter}{entry}")
                continue
            # `entry` here is a folder's own mbox file (e.g. "Fontana"); recurse
            # treating sbd_dir as the local_dir and entry as local_name
            walk_and_migrate(session, sbd_dir, entry, imap_path, delimiter, state, dry_run, only_filter)


def migrate_source(session, source_dir, dest_top_folder, delimiter, state, dry_run, only_filter):
    """
    Migrate every top-level mbox file found directly inside `source_dir` (each with
    its own optional "<name>.sbd" subfolder tree) so it lands under dest_top_folder
    in the destination mailbox. `source_dir` is typically a Thunderbird account
    folder (e.g. "Mail/mail.example.com") or a "Local Folders" directory - a
    directory holding sibling top-level mbox files directly, not a single folder's
    ".sbd" contents.
    """
    # Create the container folder itself first, even though no mail lands directly
    # in it - its children need a real parent to nest under.
    call_with_reconnect(session, ensure_folder, dest_top_folder, dry_run)

    for entry in sorted(os.listdir(source_dir)):
        if entry.endswith(SKIP_SUFFIXES):
            continue
        if entry.endswith(".sbd"):
            continue  # handled as the sibling of its base mbox file
        full = os.path.join(source_dir, entry)
        if not os.path.isfile(full):
            continue
        if _is_skipped(entry):
            log(f"Skipping folder (configured skip list): {dest_top_folder}{delimiter}{entry}")
            continue
        walk_and_migrate(session, source_dir, entry, dest_top_folder, delimiter, state, dry_run, only_filter)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Show what would happen, make no changes")
    parser.add_argument("--only", default=None, help="Only migrate folders whose destination path contains this substring")
    parser.add_argument("--reset", action="store_true", help="Ignore existing migration_state.json and start over")
    args = parser.parse_args()

    if args.reset and os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)

    for src in SOURCES:
        if not os.path.isdir(src["source_dir"]):
            log(f"ERROR: source_dir does not exist: {src['source_dir']}")
            sys.exit(1)

    state = load_state()

    session = ImapSession()
    delimiter = call_with_reconnect(session, get_delimiter)
    log(f"Connected. Folder hierarchy delimiter is: {delimiter!r}")
    if SKIP_FOLDER_NAMES:
        log(f"Skipping folders named: {sorted(SKIP_FOLDER_NAMES)}")

    for src in SOURCES:
        log(f"--- Migrating {src['source_dir']}  ->  {src['dest_top_folder']} ---")
        migrate_source(
            session, src["source_dir"], src["dest_top_folder"], delimiter, state, args.dry_run, args.only
        )

    try:
        session.imap.logout()
    except Exception:
        pass
    log("All done.")


if __name__ == "__main__":
    main()
