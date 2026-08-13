#!/usr/bin/env python3
"""
updater.py — Pulls the latest versions of this app's LOGIC files (the
parser, generator, airport lookup/resolver — the things that actually
change when a bug gets fixed) from a private GitHub repo on startup, so a
fix takes effect for everyone the next time they open the app. No exe
rebuild, no redistribution, no "please download the new version" email.

Deliberately does NOT touch the GUI shell files (invoice_portal.py,
invoice_processor.py, airport_manager.py) — a real interface change still
ships as a new build, but based on how this app has actually evolved,
that's rare; almost everything that's changed has been logic, which this
DOES cover.

Setup (one-time):
  1. Create a fine-grained Personal Access Token at
     https://github.com/settings/tokens?type=beta
       - Repository access: "Only select repositories" → this one repo
       - Permissions: Contents → Read-only (nothing else)
     This limits the blast radius if the token is ever extracted from a
     built .exe: it can only read this one repo's files, nothing else —
     not your account, not other repos, and it can't write/modify anything.
  2. Edit update_config.json (created automatically next to this file on
     first run if it doesn't exist) with your repo owner/name/branch and
     that token. This file is NOT part of the frozen exe's bundled code —
     it can be edited any time without a rebuild.

Usage (call this before importing any of the files it manages):
    import updater
    updater.sync()
    # only now import state_parser, invoice_generator, etc.
"""

import os
import sys
import json
import hashlib
import urllib.request
import urllib.error

DEFAULT_UPDATABLE_FILES = [
    "state_parser.py",
    "invoice_generator.py",
    "airport_lookup.py",
    "airport_resolver.py",
]

DEFAULT_CONFIG = {
    "owner": "YOUR_GITHUB_USERNAME_OR_ORG",
    "repo": "YOUR_REPO_NAME",
    "branch": "main",
    "token": "github_pat_PASTE_YOUR_FINE_GRAINED_TOKEN_HERE",
    "files": DEFAULT_UPDATABLE_FILES,
    "enabled": True,
}


def _app_dir() -> str:
    """Same folder this script lives in — where the config file and log
    live, and (for a non-frozen dev run) where the bundled fallback copies
    of the logic files already are."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _data_dir() -> str:
    """Same persistent, per-user, always-writable directory used for the
    airport database overrides — see airport_lookup.py's own comment on
    why this location (not a PyInstaller temp extraction folder) is the
    right choice for anything that needs to survive a restart."""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    d = os.path.join(base, "TravelWizards")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


def _config_path() -> str:
    return os.path.join(_app_dir(), "update_config.json")


def _cache_dir() -> str:
    d = os.path.join(_data_dir(), "logic_cache")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


def _log_path() -> str:
    return os.path.join(_data_dir(), "update_log.txt")


def _log(msg: str):
    line = f"{msg}\n"
    try:
        with open(_log_path(), "a", encoding="utf-8") as f:
            f.write(line)
    except OSError:
        pass
    print(f"[updater] {msg}")


def load_config() -> dict:
    """Reads update_config.json, creating it with placeholder values the
    first time this runs so there's something to edit. Never overwrites an
    existing config."""
    path = _config_path()
    if not os.path.exists(path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, indent=2)
            _log(f"Created {path} with placeholder values — edit it with "
                 "your repo owner/name/branch and token, then restart.")
        except OSError as e:
            _log(f"Could not create {path}: {e}")
        return dict(DEFAULT_CONFIG)

    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        _log(f"Could not read {path} ({e}) — using placeholder values.")
        return dict(DEFAULT_CONFIG)

    merged = dict(DEFAULT_CONFIG)
    merged.update(cfg)
    return merged


def _seed_cache_from_bundled(cache_dir: str, filenames: list):
    """First-ever run (or first run after a cache wipe): if a file isn't
    in the cache yet, seed it from whatever's bundled alongside the exe,
    so there's always a working copy even if this machine has never had
    internet access at launch time.

    filenames may be plain names ("state_parser.py") or GitHub paths that
    include a subfolder ("Invoice_Portal/state_parser.py", if that's where
    the repo actually keeps them). Either way, the LOCAL copy — both the
    bundled fallback next to the exe and the cached copy — always uses
    just the basename, since that's what has to sit directly in a folder
    on sys.path for `import state_parser` to work."""
    for entry in filenames:
        local_name = os.path.basename(entry)
        cached = os.path.join(cache_dir, local_name)
        if os.path.exists(cached):
            continue
        bundled = os.path.join(_app_dir(), local_name)
        if os.path.exists(bundled):
            try:
                with open(bundled, "rb") as src, open(cached, "wb") as dst:
                    dst.write(src.read())
                _log(f"Seeded {local_name} into cache from bundled copy.")
            except OSError as e:
                _log(f"Could not seed {local_name} from bundled copy: {e}")


def _fetch_remote_file(owner: str, repo: str, branch: str, path: str, token: str, timeout=8) -> bytes:
    """Fetches raw file bytes from GitHub's Contents API. Raises on any
    failure — caller decides how to handle that (fall back to cache)."""
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}?ref={branch}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.raw+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "TravelWizards-Updater",
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sync(quiet: bool = False) -> bool:
    """
    Checks GitHub for newer versions of every file in the config's file
    list, updates the local cache for any that changed, and puts that
    cache directory at the FRONT of sys.path so a subsequent `import
    state_parser` (etc.) picks up the fresh copy rather than any bundled
    one.

    Safe to call even with no internet, a bad token, or an unedited
    placeholder config — falls back to whatever's already cached (seeding
    from the bundled copy first, if the cache is empty), logs what
    happened, and never raises out to the caller. A network hiccup should
    never be able to stop the app from opening.

    Returns True if the app should proceed normally (cache dir is usable
    either way — this basically always returns True; it's here for
    callers that want to react to a hard failure, though there isn't
    currently a case where sync leaves the app with no usable files at
    all, short of a broken installation missing the bundled fallbacks).
    """
    cfg = load_config()
    cache_dir = _cache_dir()
    filenames = cfg.get("files") or DEFAULT_UPDATABLE_FILES

    _seed_cache_from_bundled(cache_dir, filenames)

    if not cfg.get("enabled", True):
        if not quiet:
            _log("Auto-update disabled in config — using cached/bundled files as-is.")
        sys.path.insert(0, cache_dir)
        return True

    owner, repo, branch, token = cfg["owner"], cfg["repo"], cfg["branch"], cfg["token"]
    if "YOUR_GITHUB" in owner or "PASTE_YOUR" in token:
        if not quiet:
            _log("update_config.json still has placeholder values — "
                 f"edit {_config_path()} to enable auto-update. "
                 "Using cached/bundled files for now.")
        sys.path.insert(0, cache_dir)
        return True

    updated = []
    failed = []
    for entry in filenames:
        # entry is the path GitHub actually needs (may include a subfolder,
        # e.g. "Invoice_Portal/state_parser.py") — but the file always
        # gets cached and imported under just its basename.
        local_name = os.path.basename(entry)
        try:
            remote_bytes = _fetch_remote_file(owner, repo, branch, entry, token)
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
            failed.append((local_name, str(e)))
            continue

        cached_path = os.path.join(cache_dir, local_name)
        local_bytes = b""
        if os.path.exists(cached_path):
            with open(cached_path, "rb") as f:
                local_bytes = f.read()

        if _sha256(remote_bytes) != _sha256(local_bytes):
            tmp_path = cached_path + ".tmp"
            with open(tmp_path, "wb") as f:
                f.write(remote_bytes)
            os.replace(tmp_path, cached_path)
            updated.append(local_name)

    if not quiet:
        if updated:
            _log(f"Updated: {', '.join(updated)}")
        if failed:
            for name, err in failed:
                _log(f"Could not check {name} ({err}) — using cached version.")
        if not updated and not failed:
            _log("Up to date — no changes.")

    sys.path.insert(0, cache_dir)
    return True


if __name__ == "__main__":
    sync()
    print(f"\nCache directory: {_cache_dir()}")
    print(f"Config file:      {_config_path()}")
    print(f"Log file:         {_log_path()}")