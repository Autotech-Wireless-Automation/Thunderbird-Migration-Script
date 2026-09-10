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

Before touching the network, the script does a quick local pass over every source
to count how many messages it's dealing with in total, so it can show real
percentages from the very first message rather than guessing. On a run migrating
hundreds of thousands of messages, printing one line per message would make the
console unreadable, so instead the console shows a compact, self-updating 2-line
status: what it's working on right now, and a percentage/elapsed/ETA line below it
that recalculates its throughput estimate as it goes (so the ETA adapts if a
reconnect slows things down, or a big folder speeds them back up). Every
individual message is still recorded in migration_log.txt, one line each, exactly
as before - only the screen itself gets the compact view.

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
import email
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
    """Parse one line of an IMAP LIST response into (delimiter, mailbox_name).
    The name comes back from the server in IMAP's "modified UTF-7" encoding
    (RFC 3501 5.1.3) - decode it to a normal Python str so callers can compare it
    directly against the Unicode folder names we build ourselves."""
    m = _LIST_RE.match(entry)
    if not m:
        return None, None
    delim = m.group("delim").decode(errors="replace")
    name_raw = m.group("name").decode(errors="replace").strip()
    if name_raw.startswith('"') and name_raw.endswith('"'):
        name_raw = name_raw[1:-1]
    return delim, imap_utf7_decode(name_raw)


# IMAP mailbox names are 7-bit only on the wire; anything outside ASCII (accented
# or Greek business/contact names showed up throughout George's real folder tree,
# e.g. "Shibui - Χριστόπουλος")
# has to be encoded as "modified UTF-7" (RFC 3501 5.1.3: like UTF-7, but "&" instead
# of "+" as the shift character, no "=" padding, and "/" replaced by "," inside a
# shifted run) before it's sent, and decoded back out of anything the server sends.
# imaplib itself has no support for this - it just does bytes(name, "ascii") and
# throws UnicodeEncodeError the moment a name isn't pure ASCII, which is exactly
# the crash a real (non-English) mailbox tree hits on the very first CREATE.

def imap_utf7_encode(name):
    """Encode a Unicode mailbox/folder name (which may contain a '/' hierarchy
    delimiter - that's fine, '/' is ASCII and passes through unchanged) into the
    bytes IMAP expects on the wire."""
    if isinstance(name, bytes):
        return name
    out = bytearray()
    i, n = 0, len(name)
    while i < n:
        ch = name[i]
        if ch == "&":
            out += b"&-"
            i += 1
            continue
        if " " <= ch <= "~":  # printable ASCII: passes through untouched
            out.append(ord(ch))
            i += 1
            continue
        j = i
        while j < n and not (" " <= name[j] <= "~"):
            j += 1
        run = name[i:j]
        b64 = base64.b64encode(run.encode("utf-16-be")).decode("ascii")
        b64 = b64.rstrip("=").replace("/", ",")
        out += b"&" + b64.encode("ascii") + b"-"
        i = j
    return bytes(out)


def imap_utf7_decode(name):
    """Inverse of imap_utf7_encode - turn a name the server sent (str or bytes,
    still in modified UTF-7) back into a normal Unicode str."""
    if isinstance(name, bytes):
        name = name.decode("ascii", errors="replace")
    out = []
    i, n = 0, len(name)
    while i < n:
        ch = name[i]
        if ch != "&":
            out.append(ch)
            i += 1
            continue
        if i + 1 < n and name[i + 1] == "-":
            out.append("&")
            i += 2
            continue
        end = name.find("-", i + 1)
        if end == -1:
            end = n
        b64 = name[i + 1:end].replace(",", "/")
        b64 += "=" * (-len(b64) % 4)
        try:
            out.append(base64.b64decode(b64).decode("utf-16-be"))
        except Exception:
            pass  # malformed run - drop it rather than raise on a cosmetic decode
        i = end + 1
    return "".join(out)


def imap_quote(name):
    """Modified-UTF-7-encode a mailbox name and wrap it in double quotes, as
    bytes ready to hand straight to imaplib (bypassing its ascii-only default
    encoder, which is what crashes on any non-ASCII folder name)."""
    return b'"' + imap_utf7_encode(name) + b'"'

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

# Exact (case-insensitive) filenames Thunderbird drops directly inside an account
# folder that are also NOT mail data - account-level settings/state, not a folder.
# Seen in the wild: msgFilterRules.dat (per-account filter rules), popstate.dat
# (POP3 UIDL/state tracking), filterlog.html (optional filter activity log).
NON_MAIL_FILENAMES = {"msgfilterrules.dat", "popstate.dat", "filterlog.html"}


def _is_metadata_file(name):
    return name.lower() in NON_MAIL_FILENAMES

APPEND_RETRY_ATTEMPTS = 5
APPEND_RETRY_BASE_DELAY = 2.0   # seconds, doubles each retry (exponential backoff)
APPEND_PACING_DELAY = 0.15      # small delay between appends to be gentle on the server

MAX_RECONNECTS = 3              # how many times to reconnect+retry a single IMAP op
RECONNECT_PAUSE = 3.0           # seconds to wait before reconnecting

# ============================================================================


_PROGRESS = None  # set once main() has sized up the run; see ProgressTracker


def log(msg):
    """Console + file. For real events (signed in, folder created/skipped,
    reconnects, per-folder summaries) - infrequent enough not to spam the screen.
    Plays nicely with the live status block, if one is active: temporarily clears
    it, prints this line above where it was, then redraws it underneath."""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    if _PROGRESS is not None:
        _PROGRESS.clear()
    print(line, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    if _PROGRESS is not None:
        _PROGRESS.repaint()


def log_file(msg):
    """File only - the detailed per-message trail. Kept in full in
    migration_log.txt for later audit/debugging, but never printed to the screen:
    at hundreds of thousands of messages that would be unreadable, which is what
    the live status block (see ProgressTracker) is for instead."""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


class ProgressTracker:
    """Renders a compact, in-place 2-line status instead of scrolling one line per
    message: line 1 is what it's working on right now, line 2 is percent complete
    (to 2 decimals), elapsed time, and a self-calibrating ETA. "Self-calibrating"
    means the throughput estimate is recomputed from actual recent progress (the
    last WINDOW_SECONDS), not a fixed guess - so it adapts if a reconnect slows
    things down or a big folder speeds them back up, falling back to the
    since-the-start average whenever there isn't enough recent data yet (e.g. right
    after startup, or right after a reconnect pause)."""

    WINDOW_SECONDS = 30.0        # only look at the last N seconds of throughput for the ETA
    REPAINT_MIN_INTERVAL = 0.1  # seconds - don't repaint faster than this

    def __init__(self, grand_total, already_done=0):
        self.grand_total = grand_total
        self.done = already_done
        self.start_time = time.time()
        self._samples = [(self.start_time, already_done)]
        self._last_repaint = 0.0
        self._active = False   # whether the 2-line status block is currently on screen
        self._last_label = ""

    def _rate(self):
        now = time.time()
        cutoff = now - self.WINDOW_SECONDS
        trimmed = [(t, n) for (t, n) in self._samples if t >= cutoff]
        self._samples = trimmed or self._samples[-1:]
        t0, n0 = self._samples[0]
        t1, n1 = self._samples[-1]
        if t1 > t0 and n1 > n0:
            recent_rate = (n1 - n0) / (t1 - t0)
            if recent_rate > 0:
                return recent_rate
        # Not enough recent data (just started, or mid-reconnect-pause) - fall back
        # to the overall average so far rather than showing a stalled/zero ETA.
        elapsed = now - self.start_time
        return (self.done / elapsed) if elapsed > 0 and self.done > 0 else 0.0

    @staticmethod
    def _fmt_duration(seconds):
        if seconds is None:
            return "calculating..."
        seconds = max(0, int(seconds))
        h, rem = divmod(seconds, 3600)
        m, s = divmod(rem, 60)
        if h:
            return f"{h}h {m:02d}m {s:02d}s"
        if m:
            return f"{m}m {s:02d}s"
        return f"{s}s"

    def advance(self, n=1):
        self.done += n
        self._samples.append((time.time(), self.done))

    def _lines(self, label):
        now = time.time()
        pct = (self.done / self.grand_total * 100.0) if self.grand_total else 100.0
        elapsed = now - self.start_time
        rate = self._rate()
        remaining = max(0, self.grand_total - self.done)
        eta_seconds = (remaining / rate) if rate > 0 else None
        finish_clock = (
            time.strftime("%H:%M:%S", time.localtime(now + eta_seconds))
            if eta_seconds is not None else "--:--:--"
        )
        line1 = f"Working on: {label}"
        line2 = (
            f"{pct:6.2f}%  ({self.done:,}/{self.grand_total:,} messages)  "
            f"elapsed {self._fmt_duration(elapsed)}  "
            f"ETA {self._fmt_duration(eta_seconds)} (finish ~{finish_clock})"
        )
        return line1[:140], line2[:140]

    def render(self, label, force=False):
        now = time.time()
        self._last_label = label
        if not force and (now - self._last_repaint) < self.REPAINT_MIN_INTERVAL:
            return
        self._last_repaint = now
        line1, line2 = self._lines(label)
        if self._active:
            sys.stdout.write("\x1b[2A")  # cursor up 2 lines, back to the start of line1
        sys.stdout.write("\x1b[2K" + line1 + "\n")
        sys.stdout.write("\x1b[2K" + line2 + "\n")
        sys.stdout.flush()
        self._active = True

    def clear(self):
        """Temporarily remove the status block so a normal scrolling log line can
        be printed where it was; call repaint() afterwards to put it back."""
        if self._active:
            sys.stdout.write("\x1b[2A")  # up to the start of line1
            sys.stdout.write("\x1b[0J")  # erase everything from there to end of screen
            sys.stdout.flush()
            self._active = False

    def repaint(self):
        if self._last_label:
            self.render(self._last_label, force=True)

    def finish(self):
        if self._last_label:
            self.render(self._last_label, force=True)
        if self._active:
            print()  # leave a blank line so later log() calls scroll normally below
            self._active = False


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
        log_file(f"[dry-run] would CREATE folder: {path}")
        _existing_folders_cache.add(path)
        return
    typ, data = imap.create(imap_quote(path))
    if typ != "OK":
        raise RuntimeError(f"Could not create folder {path}: {data}")
    _existing_folders_cache.add(path)
    log_file(f"Created folder: {path}")


class PermanentAppendError(Exception):
    """The server flatly rejected this one specific message - most commonly
    Exchange Online's IMAP APPEND per-message size ceiling, which is separate
    from (and can be lower than) the mailbox's own send/receive size limits, and
    can't be raised from this side. Retrying the identical APPEND will never
    succeed, so migrate_one_folder skips just this message and keeps going
    instead of letting the whole multi-hour run die over one oversized email."""


# Case-insensitive substrings in an IMAP NO/BAD response that mean "this exact
# message will never go through, no matter how many times you retry it" - fail
# fast on these rather than burning through APPEND_RETRY_ATTEMPTS of backoff.
_PERMANENT_APPEND_MARKERS = (
    "maximum size of appendable message",
    "message too large",
    "size limit",
    "over quota",
    "quota exceeded",
)


def _is_permanent_append_error(err):
    text = str(err).lower()
    return any(marker in text for marker in _PERMANENT_APPEND_MARKERS)


def append_message(imap, folder_path, raw_bytes, internaldate):
    last_err = None
    for attempt in range(1, APPEND_RETRY_ATTEMPTS + 1):
        try:
            typ, data = imap.append(imap_quote(folder_path), None, internaldate, raw_bytes)
            if typ == "OK":
                return
            last_err = data
        except imaplib.IMAP4.abort:
            # connection dropped - bubble up so caller can reconnect
            raise
        except Exception as e:
            last_err = e

        if _is_permanent_append_error(last_err):
            raise PermanentAppendError(str(last_err))

        if attempt < APPEND_RETRY_ATTEMPTS:
            delay = APPEND_RETRY_BASE_DELAY * (2 ** (attempt - 1))
            # A retry is a real, occasionally-useful-to-notice event, but on a run
            # this long it can also repeat often enough to spam the console - keep
            # it in the file; a stalled elapsed-time on the live status line is
            # the on-screen tell that something's being retried.
            log_file(f"  append retry {attempt}/{APPEND_RETRY_ATTEMPTS} after error: {last_err} (sleeping {delay:.0f}s)")
            time.sleep(delay)

    # Exhausted every retry on something that wasn't recognized as an immediately-
    # permanent rejection either - still treat it as "skip this one message and
    # move on" rather than as a reason to kill the whole run.
    raise PermanentAppendError(f"gave up after {APPEND_RETRY_ATTEMPTS} attempts: {last_err}")


# ----------------------------- state tracking ----------------------------------

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"folders_done": {}, "folder_progress": {}}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def maybe_save_state(state, dry_run):
    """A --dry-run must never persist migration_state.json - it hasn't actually
    created any folders or appended any messages, so marking anything "done" would
    make a *real* run afterwards think that work is finished and silently skip it.
    Real runs always save; dry runs never do (state is still updated in memory for
    the duration of this process, only so --only / repeated-folder edge cases and
    the live progress numbers behave sensibly within the one dry-run pass)."""
    if not dry_run:
        save_state(state)


# ----------------------------- mbox walking ------------------------------------

def internaldate_for(raw_bytes):
    try:
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


SKIPPED_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skipped_messages.txt")


def _describe_message(raw_bytes):
    """Best-effort Subject/Date/size for a message that couldn't be migrated, so
    it's findable later without digging through the raw mbox file."""
    try:
        msg = email.message_from_bytes(raw_bytes)
        subject = msg.get("Subject", "(no subject)")
        date_hdr = msg.get("Date", "(no date)")
    except Exception:
        subject, date_hdr = "(unreadable)", "(unreadable)"
    return subject, date_hdr, len(raw_bytes)


def record_skipped_message(state, imap_path, index, total, raw_bytes, error):
    """A message the server permanently rejected (see PermanentAppendError) -
    logged loudly (this is real data not making it across, unlike the routine
    per-folder bookkeeping) and appended to skipped_messages.txt in one line per
    message so everything that didn't make it can be reviewed/handled by hand
    afterwards, without having to dig through the full migration_log.txt."""
    subject, date_hdr, size = _describe_message(raw_bytes)
    reason = str(error).strip()
    log(
        f"SKIPPED message {index + 1}/{total} in {imap_path} (server permanently rejected it): {reason} "
        f"[Subject: {subject!r}  Date: {date_hdr}  Size: {size:,} bytes]"
    )
    with open(SKIPPED_FILE, "a", encoding="utf-8") as f:
        f.write(
            f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {imap_path} #{index + 1}/{total} "
            f"size={size} date={date_hdr!r} subject={subject!r} reason={reason!r}\n"
        )
    state["skipped_count"] = state.get("skipped_count", 0) + 1


def migrate_one_folder(session, local_folder_file, imap_path, state, dry_run):
    """Append all messages in a single mbox file to imap_path, resuming if needed.

    Every per-folder event here (already-done, folder created, message counts,
    finished) goes to log_file() only, not log() - on a tree with thousands of
    small folders these fire constantly, which is exactly the console spam the
    live status line replaces. Only the status line (line 1: current folder and
    position; line 2: percent/elapsed/ETA) is meant to visibly change per folder;
    everything else stays in migration_log.txt for later review."""
    if state["folders_done"].get(imap_path):
        log_file(f"Skipping (already done): {imap_path}")
        return

    # Always create the folder itself, even if it holds no direct messages -
    # its children (if any) need a real parent to nest under.
    call_with_reconnect(session, ensure_folder, imap_path, dry_run)
    if _PROGRESS is not None:
        _PROGRESS.render(imap_path)

    if not os.path.exists(local_folder_file) or os.path.getsize(local_folder_file) == 0:
        state["folders_done"][imap_path] = True
        maybe_save_state(state, dry_run)
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
    log_file(f"{imap_path}: {total} message(s) in source, resuming at {start_index}")
    for i, key in enumerate(keys):
        if i < start_index:
            continue
        raw = box.get_bytes(key)
        idate = internaldate_for(raw)
        if dry_run:
            log_file(f"[dry-run] would APPEND message {i + 1}/{total} to {imap_path}")
        else:
            try:
                call_with_reconnect(session, append_message, imap_path, raw, idate)
                log_file(f"APPENDed message {i + 1}/{total} to {imap_path}")
            except PermanentAppendError as e:
                # The server will never accept this exact message (most often:
                # Exchange Online's IMAP APPEND size ceiling) - record it and move
                # on to the next message rather than losing the whole multi-hour
                # run over one email.
                record_skipped_message(state, imap_path, i, total, raw, e)
            time.sleep(APPEND_PACING_DELAY)
        state["folder_progress"][imap_path] = i + 1
        if _PROGRESS is not None:
            _PROGRESS.advance(1)
            _PROGRESS.render(f"{imap_path}  [{i + 1:,}/{total:,}]")
        if (i + 1) % 25 == 0:
            maybe_save_state(state, dry_run)

    state["folders_done"][imap_path] = True
    maybe_save_state(state, dry_run)
    log_file(f"Finished: {imap_path} ({total} messages)")


def _is_skipped(name):
    return name.lower() in SKIP_FOLDER_NAMES


def iter_folder_tree(local_dir, local_name, imap_parent_path, delimiter, only_filter, announce):
    """
    Yields (imap_path, mbox_file) for `local_name` itself and everything nested
    under its "<local_name>.sbd" subfolder tree, honoring --only and
    skip_folder_names exactly as the real run will. Pure filesystem walk - no
    IMAP, no state - so the planning pass (which sizes up the progress bar) and
    the real migration pass (which actually appends messages) can share this one
    definition of "what gets migrated" instead of two that could drift apart.

    local_dir: directory containing `local_name` (mbox file) and `local_name.sbd`
    (subfolder dir), if any.
    """
    imap_path = f"{imap_parent_path}{delimiter}{local_name}" if imap_parent_path else local_name

    mbox_file = os.path.join(local_dir, local_name)
    sbd_dir = os.path.join(local_dir, local_name + ".sbd")

    do_this_folder = (not only_filter) or (only_filter in imap_path)

    if do_this_folder:
        yield (imap_path, mbox_file)

    if os.path.isdir(sbd_dir):
        for entry in sorted(os.listdir(sbd_dir)):
            if entry.endswith(SKIP_SUFFIXES) or _is_metadata_file(entry):
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
                if announce:
                    log_file(f"Skipping folder (configured skip list): {imap_path}{delimiter}{entry}")
                continue
            # `entry` here is a folder's own mbox file (e.g. "Fontana"); recurse
            # treating sbd_dir as the local_dir and entry as local_name
            yield from iter_folder_tree(sbd_dir, entry, imap_path, delimiter, only_filter, announce)


def iter_source_tree(source_dir, dest_top_folder, delimiter, only_filter, announce):
    """
    Yields (imap_path, mbox_file) for every top-level mbox file found directly
    inside `source_dir` (each with its own optional "<name>.sbd" subfolder tree),
    landing under dest_top_folder. `source_dir` is typically a Thunderbird account
    folder (e.g. "Mail/mail.example.com") or a "Local Folders" directory - a
    directory holding sibling top-level mbox files directly, not a single folder's
    ".sbd" contents. Does not include dest_top_folder itself (the container has no
    messages of its own; see migrate_source for creating it).
    """
    for entry in sorted(os.listdir(source_dir)):
        if entry.endswith(SKIP_SUFFIXES) or _is_metadata_file(entry):
            continue
        if entry.endswith(".sbd"):
            continue  # handled as the sibling of its base mbox file
        full = os.path.join(source_dir, entry)
        if not os.path.isfile(full):
            continue
        if _is_skipped(entry):
            if announce:
                log_file(f"Skipping folder (configured skip list): {dest_top_folder}{delimiter}{entry}")
            continue
        yield from iter_folder_tree(source_dir, entry, dest_top_folder, delimiter, only_filter, announce)


def _count_messages(mbox_file):
    if not os.path.exists(mbox_file) or os.path.getsize(mbox_file) == 0:
        return 0
    box = mailbox.mbox(mbox_file, factory=None, create=False)
    return len(box.keys())


def plan_totals(sources, delimiter, only_filter, state):
    """One quiet, local-only pass over every source (no network, no announcing of
    skips - the real run does that) to total up exactly how many messages this run
    is dealing with, and how many of those are already done from a previous run -
    so the progress bar can show accurate numbers from message #1 instead of
    growing its denominator as it goes."""
    grand_total = 0
    already_done = 0
    for src in sources:
        for imap_path, mbox_file in iter_source_tree(
            src["source_dir"], src["dest_top_folder"], delimiter, only_filter, announce=False
        ):
            total = _count_messages(mbox_file)
            grand_total += total
            if state["folders_done"].get(imap_path):
                already_done += total
            else:
                already_done += min(state["folder_progress"].get(imap_path, 0), total)
    return grand_total, already_done


def migrate_source(session, source_dir, dest_top_folder, delimiter, state, dry_run, only_filter):
    """
    Migrate every top-level mbox file found directly inside `source_dir` (each with
    its own optional "<name>.sbd" subfolder tree) so it lands under dest_top_folder
    in the destination mailbox.
    """
    # Create the container folder itself first, even though no mail lands directly
    # in it - its children need a real parent to nest under.
    call_with_reconnect(session, ensure_folder, dest_top_folder, dry_run)

    for imap_path, mbox_file in iter_source_tree(source_dir, dest_top_folder, delimiter, only_filter, announce=True):
        migrate_one_folder(session, mbox_file, imap_path, state, dry_run)


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

    log("Scanning source folders to size up the run (local disk only, no network yet)...")
    grand_total, already_done = plan_totals(SOURCES, delimiter, args.only, state)
    global _PROGRESS
    _PROGRESS = ProgressTracker(grand_total, already_done)
    log(
        f"Found {grand_total:,} message(s) across {len(SOURCES)} source(s) to account for "
        f"({already_done:,} already done from a previous run)."
    )

    for src in SOURCES:
        log(f"--- Migrating {src['source_dir']}  ->  {src['dest_top_folder']} ---")
        migrate_source(
            session, src["source_dir"], src["dest_top_folder"], delimiter, state, args.dry_run, args.only
        )

    _PROGRESS.finish()

    try:
        session.imap.logout()
    except Exception:
        pass

    if state.get("skipped_count"):
        log(
            f"NOTE: {state['skipped_count']} message(s) could not be migrated - the server "
            f"permanently rejected them (most likely Exchange Online's IMAP APPEND size ceiling). "
            f"See skipped_messages.txt for exactly which ones (folder, subject, date, size)."
        )
    log("All done.")


if __name__ == "__main__":
    main()
